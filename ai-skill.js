/**
 * AI Skill 调用模块
 * 通过本地 API 服务调用豆包
 */

// ==================== 配置 ====================

// 本地 API 服务地址（Skill + 模型桥接服务）
const DEFAULT_LOCAL_AI_API_URL = "http://127.0.0.1:8765/analyze";

function getAnalyzeApiUrl() {
    try {
        if (typeof window !== "undefined" && window.PEVCRuntimeConfig && window.PEVCRuntimeConfig.buildApiUrl) {
            return window.PEVCRuntimeConfig.buildApiUrl("/analyze");
        }
    } catch (_err) {}
    return DEFAULT_LOCAL_AI_API_URL;
}

// 文档切片大小（使用独特名称避免与 taskpane.js 中的 CHUNK_SIZE 冲突）
const AI_SKILL_CHUNK_SIZE = 1000;  // 字符（按段落聚合）

// ==================== 文档切片 ====================

/**
 * 按行切分文档，聚合到接近 chunkSize
 * 
 * 策略：
 * 1. 按单个 \n 分行
 * 2. 聚合相邻行到接近 chunkSize
 * 3. 单行超过 chunkSize 则独立成块（不拆行）
 * 
 * @param {string} text - 文档文本
 * @param {number} chunkSize - 目标切片大小（字符）
 * @returns {array} 文档片段数组
 */
function chunkDocument(text, chunkSize = AI_SKILL_CHUNK_SIZE) {
    if (!text || text.length <= chunkSize) {
        return [text];
    }
    
    // 1. 统一换行符格式，然后按换行分割
    // 兼容 Windows (\r\n)、Mac (\r)、Unix (\n) 以及中文句号分段
    const normalizedText = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    let lines = normalizedText.split(/\n/).map(l => l.trim()).filter(l => l);
    
    // 如果分割后只有 1 行且超过 chunkSize，尝试按中文句号分割
    if (lines.length === 1 && lines[0].length > chunkSize) {
        console.log('[AI Skill] 换行分割无效，尝试按句号分割');
        lines = text.split(/[。！？；]/).map(l => l.trim()).filter(l => l);
    }
    
    if (!lines || lines.length === 0) {
        return [text];
    }
    
    // 2. 聚合行到接近 chunkSize
    const chunks = [];
    let currentChunkParts = [];
    let currentLength = 0;
    
    for (const line of lines) {
        const lineLen = line.length;
        
        // 如果单行就超过 chunkSize，独立成块
        if (lineLen >= chunkSize) {
            // 先保存当前聚合的块（如果有）
            if (currentChunkParts.length > 0) {
                chunks.push(currentChunkParts.join("\n"));
                currentChunkParts = [];
                currentLength = 0;
            }
            // 单行独立成块
            chunks.push(line);
            continue;
        }
        
        // 如果加上当前行会超过 chunkSize，先保存前面的块
        if (currentLength > 0 && currentLength + lineLen + 1 > chunkSize) {  // +1 for \n
            chunks.push(currentChunkParts.join("\n"));
            currentChunkParts = [line];
            currentLength = lineLen;
        } else {
            // 聚合到当前块
            currentChunkParts.push(line);
            currentLength += lineLen + 1;  // +1 for \n separator
        }
    }
    
    // 保存最后一块
    if (currentChunkParts.length > 0) {
        chunks.push(currentChunkParts.join("\n"));
    }
    
    console.log(`[AI Skill] 文档切分为 ${chunks.length} 个片段（目标大小: ${chunkSize}）`);
    chunks.forEach((chunk, i) => {
        console.log(`  片段 ${i+1}: ${chunk.length} 字符`);
    });
    
    return chunks;
}

// ==================== AI 调用 ====================

/**
 * 加载本地 Skill 提示词文档
 * 兼容历史文件名 Skill.md 与标准文件名 SKILL.md
 * @returns {Promise<string>} Skill 文本
 */
async function loadSkillPrompt() {
    const candidates = [
        './contract-variable-skill/SKILL.md',
        './contract-variable-skill/Skill.md'
    ];

    let lastError = null;
    for (const path of candidates) {
        try {
            const response = await fetch(path, { cache: 'no-store' });
            if (!response.ok) {
                lastError = new Error(`读取失败: ${path} (${response.status})`);
                continue;
            }

            const text = await response.text();
            if (text && text.trim()) {
                return text;
            }

            lastError = new Error(`文件为空: ${path}`);
        } catch (error) {
            lastError = error;
        }
    }

    throw new Error(`无法加载 Skill 提示词: ${lastError ? lastError.message : '未知错误'}`);
}

/**
 * 调用 AI API（单个片段）
 * @param {string} documentText - 文档文本
 * @param {number} chunkIndex - 当前片段索引
 * @param {number} totalChunks - 总片段数
 * @returns {Promise<object>} AI 返回的 JSON 对象
 */
async function callAIWithSkill(documentText, chunkIndex = 0, totalChunks = 1, chunkSize = AI_SKILL_CHUNK_SIZE) {
    console.log(`[AI Skill] 开始调用 AI... (分块 ${chunkIndex + 1}/${totalChunks})`);
    console.log(`[AI Skill] 文档长度: ${documentText.length} 字符`);
    
    try {
        const response = await fetch(getAnalyzeApiUrl(), {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                text: documentText,
                chunk_size: chunkSize
            })
        });
        
        if (!response.ok) {
            const errText = await response.text();
            throw new Error(`本地 AI 接口失败: ${response.status} - ${errText}`);
        }
        
        const jsonData = await response.json();
        console.log("[AI Skill] 本地接口返回:", jsonData);

        if (!window.AIParser) {
            console.warn('[AI Skill] AIParser 未加载，跳过验证');
            return jsonData;
        }

        const validation = window.AIParser.validateAIOutput(jsonData);
        if (!validation.valid) {
            console.error('[AI Skill] AI 输出验证失败:', validation.errors);
            window.AIParser.logUnknownFormats(jsonData);
        } else {
            console.log('[AI Skill] AI 输出验证通过');
        }
        if (validation.warnings && validation.warnings.length > 0) {
            console.warn('[AI Skill] AI 输出警告:', validation.warnings);
        }

        return jsonData;
    } catch (error) {
        console.error('[AI Skill] API 调用失败:', error);
        throw error;
    }
}

/**
 * 分析完整文档（支持自动切片和进度回调）
 * @param {string} fullText - 完整文档文本
 * @param {function} onProgress - 进度回调函数 (current, total, status)
 * @returns {Promise<object>} 合并后的 AI 输出
 */
async function analyzeDocument(fullText, onProgress) {
    console.log('[AI Skill] 开始分析文档...');
    
    // 1. 文档切片
    const chunks = chunkDocument(fullText, AI_SKILL_CHUNK_SIZE);
    const totalChunks = chunks.length;
    const allVariables = [];
    let successChunks = 0;
    let firstChunkError = null;
    
    if (onProgress) onProgress(0, totalChunks, "开始分析文档...");
    
    // 2. 逐个片段处理
    for (let i = 0; i < totalChunks; i++) {
        const chunk = chunks[i];
        const currentNum = i + 1;
        
        if (onProgress) {
            onProgress(i, totalChunks, `正在识别第 ${currentNum}/${totalChunks} 个片段...`);
        }
        
        try {
            const result = await callAIWithSkill(chunk, i, totalChunks, AI_SKILL_CHUNK_SIZE);
            if (result && result.variables && Array.isArray(result.variables)) {
                allVariables.push(...result.variables);
            }
            successChunks += 1;
        } catch (error) {
            console.error(`[AI Skill] 片段 ${currentNum} 处理失败:`, error);
            if (!firstChunkError) firstChunkError = error;
            // 如果只有一块且失败，则抛出错误；多块时可选择跳过或继续
            if (totalChunks === 1) throw error;
        }
    }

    if (successChunks === 0 && firstChunkError) {
        throw firstChunkError;
    }
    
    if (onProgress) onProgress(totalChunks, totalChunks, "解析完成，正在整合结果...");
    
    // 3. 结果整合与去重
    const deduplicated = deduplicateVariables(allVariables);
    
    return {
        variables: deduplicated
    };
}

/**
 * 去重变量（根据 tag）
 * @param {array} variables - 变量数组
 * @returns {array} 去重后的变量数组
 */
function deduplicateVariables(variables) {
    const seen = new Map();
    const unique = [];
    
    variables.forEach(variable => {
        if (!variable.tag) {
            unique.push(variable);
            return;
        }
        
        if (seen.has(variable.tag)) {
            // 已存在，根据 confidence 和 layer 决定是否替换
            const existing = seen.get(variable.tag);
            
            // 优先保留 confidence 更高的
            const confidencePriority = { high: 3, medium: 2, low: 1 };
            const existingConfidence = confidencePriority[existing.confidence || 'medium'];
            const newConfidence = confidencePriority[variable.confidence || 'medium'];
            
            if (newConfidence > existingConfidence || 
                (newConfidence === existingConfidence && (variable.layer || 1) < (existing.layer || 1))) {
                // 替换
                const index = unique.indexOf(existing);
                unique[index] = variable;
                seen.set(variable.tag, variable);
            }
        } else {
            seen.set(variable.tag, variable);
            unique.push(variable);
        }
    });
    
    const deduped = unique.length;
    const original = variables.length;
    if (deduped < original) {
        console.log(`[AI Skill] 去重：${original} → ${deduped} (移除 ${original - deduped} 个重复项)`);
    }
    
    return unique;
}

// ==================== 导出 ====================

// 兼容浏览器环境
if (typeof window !== 'undefined') {
    window.AISkill = {
        analyzeDocument,
        callAIWithSkill,
        chunkDocument,
        deduplicateVariables,
        loadSkillPrompt
    };
}

// 兼容 Node.js 环境
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        analyzeDocument,
        callAIWithSkill,
        chunkDocument,
        deduplicateVariables,
        loadSkillPrompt
    };
}
