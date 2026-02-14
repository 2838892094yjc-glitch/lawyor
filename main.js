/**
 * PEVC合同助手 - WPS版 主入口文件
 * Ribbon 回调函数 & 初始化
 */

// =====================================================================
// Ribbon 回调函数
// =====================================================================
const PEVC_BOOTSTRAP_STORAGE_KEY = "pevc_bootstrap_config_v1";
const PEVC_RUNTIME_STORAGE_KEY = "pevc_runtime_config_v1";
const PEVC_DEFAULT_BOOTSTRAP = {
    // 建议上线后替换为你的云端 manifest 地址。
    manifestUrl: "",
    // 如果你希望强制固定远端 taskpane，可直接写这里。
    taskpaneUrl: "",
    requestTimeoutMs: 2500,
    preferRemote: true
};

function readJsonStorage(key) {
    try {
        const raw = localStorage.getItem(key);
        return raw ? JSON.parse(raw) : null;
    } catch (_err) {
        return null;
    }
}

function writeJsonStorage(key, value) {
    try {
        localStorage.setItem(key, JSON.stringify(value));
        return true;
    } catch (_err) {
        return false;
    }
}

function getBootstrapConfig() {
    const saved = readJsonStorage(PEVC_BOOTSTRAP_STORAGE_KEY) || {};
    const runtime = readJsonStorage(PEVC_RUNTIME_STORAGE_KEY) || {};
    return {
        manifestUrl: saved.manifestUrl || runtime.manifestUrl || PEVC_DEFAULT_BOOTSTRAP.manifestUrl,
        taskpaneUrl: saved.taskpaneUrl || runtime.taskpaneUrl || PEVC_DEFAULT_BOOTSTRAP.taskpaneUrl,
        preferRemote: saved.preferRemote !== undefined ? !!saved.preferRemote : !!PEVC_DEFAULT_BOOTSTRAP.preferRemote,
        requestTimeoutMs: typeof saved.requestTimeoutMs === "number" ? saved.requestTimeoutMs : PEVC_DEFAULT_BOOTSTRAP.requestTimeoutMs
    };
}

function appendBootstrapParams(rawUrl, cfg) {
    if (!rawUrl) return rawUrl;

    try {
        const baseHref = (typeof window !== "undefined" && window.location && window.location.href)
            ? window.location.href
            : undefined;
        const urlObj = baseHref ? new URL(rawUrl, baseHref) : new URL(rawUrl);

        if (cfg && cfg.manifestUrl) {
            urlObj.searchParams.set("pevc_manifest", cfg.manifestUrl);
        }
        if (cfg && cfg.requestTimeoutMs) {
            urlObj.searchParams.set("pevc_manifest_timeout_ms", String(cfg.requestTimeoutMs));
        }

        return urlObj.toString();
    } catch (_err) {
        return rawUrl;
    }
}

function extractTaskpaneUrl(payload) {
    if (!payload || typeof payload !== "object") return "";
    return payload.taskpaneUrl || (payload.client && payload.client.taskpaneUrl) || "";
}

function syncRuntimeBootstrapConfig(cfg, payload) {
    const current = readJsonStorage(PEVC_RUNTIME_STORAGE_KEY) || {};
    const next = Object.assign({}, current);

    if (cfg && cfg.manifestUrl) next.manifestUrl = cfg.manifestUrl;
    if (cfg && cfg.taskpaneUrl) next.taskpaneUrl = cfg.taskpaneUrl;

    if (payload && typeof payload === "object") {
        const taskpaneUrl = extractTaskpaneUrl(payload);
        const apiBaseUrl = payload.apiBaseUrl
            || (payload.services && payload.services.apiBaseUrl)
            || (payload.services && payload.services.aiBaseUrl)
            || "";

        if (taskpaneUrl) next.taskpaneUrl = taskpaneUrl;
        if (apiBaseUrl) next.apiBaseUrl = apiBaseUrl;
        if (payload.version || payload.latestVersion) {
            next.version = payload.version || payload.latestVersion;
        }
        if (payload.autoRefreshMs || (payload.client && payload.client.autoRefreshMs)) {
            next.autoRefreshMs = payload.autoRefreshMs || payload.client.autoRefreshMs;
        }
    }

    writeJsonStorage(PEVC_RUNTIME_STORAGE_KEY, next);
}

function fetchManifestByXhr(manifestUrl, timeoutMs) {
    return new Promise((resolve) => {
        try {
            if (typeof XMLHttpRequest === "undefined") return resolve(null);
            const xhr = new XMLHttpRequest();
            xhr.open("GET", manifestUrl, true);
            xhr.timeout = timeoutMs || 2500;
            xhr.onreadystatechange = function () {
                if (xhr.readyState !== 4) return;
                if (xhr.status < 200 || xhr.status >= 300) return resolve(null);
                try {
                    resolve(JSON.parse(xhr.responseText || "{}"));
                } catch (_parseErr) {
                    resolve(null);
                }
            };
            xhr.onerror = function () { resolve(null); };
            xhr.ontimeout = function () { resolve(null); };
            xhr.send();
        } catch (_err) {
            resolve(null);
        }
    });
}

async function fetchRemoteManifest(manifestUrl, timeoutMs) {
    if (!manifestUrl) return null;

    if (typeof fetch !== "function") {
        return fetchManifestByXhr(manifestUrl, timeoutMs);
    }

    const controller = typeof AbortController === "function" ? new AbortController() : null;
    const timer = setTimeout(() => {
        if (controller) controller.abort();
    }, timeoutMs || 2500);

    try {
        const response = await fetch(manifestUrl, {
            cache: "no-store",
            signal: controller ? controller.signal : undefined
        });
        if (!response.ok) return null;
        return await response.json();
    } catch (_err) {
        return fetchManifestByXhr(manifestUrl, timeoutMs);
    } finally {
        clearTimeout(timer);
    }
}

async function resolveDialogUrl() {
    const fallbackRawUrl = (() => {
        try {
            if (typeof window !== "undefined" && window.location && window.location.href) {
                return new URL("./taskpane.html", window.location.href).href;
            }
        } catch (_err) {}
        return "./taskpane.html";
    })();

    const cfg = getBootstrapConfig();

    if (!cfg.preferRemote) {
        syncRuntimeBootstrapConfig(cfg, null);
        return appendBootstrapParams(fallbackRawUrl, cfg);
    }

    if (cfg.taskpaneUrl) {
        syncRuntimeBootstrapConfig(cfg, null);
        return appendBootstrapParams(cfg.taskpaneUrl, cfg);
    }

    const manifestPayload = await fetchRemoteManifest(cfg.manifestUrl, cfg.requestTimeoutMs);
    syncRuntimeBootstrapConfig(cfg, manifestPayload);

    const remote = extractTaskpaneUrl(manifestPayload);
    return appendBootstrapParams(remote || fallbackRawUrl, cfg);
}

function OnAddinLoad(ribbonUI) {
    console.log("[WPS] PEVC合同助手已加载");
    window.ribbon = ribbonUI;
    return true;
}

function OnShowDialog() {
    console.log("[WPS] OnShowDialog called");
    resolveDialogUrl()
        .then((dialogUrl) => {
            console.log("[WPS] ShowDialog URL:", dialogUrl);
            wps.ShowDialog(dialogUrl, "PEVC合同助手", 420, 700, false);
        })
        .catch((err) => {
            console.error("[WPS] resolve dialog url failed:", err);
            try {
                wps.ShowDialog("./taskpane.html", "PEVC合同助手", 420, 700, false);
            } catch (showErr) {
                console.error("[WPS] ShowDialog error:", showErr);
            }
        });
    return true;
}

function OnApiTest() {
    console.log("[WPS] OnApiTest called");
    runApiTests();
    return true;
}

function OnApplyForm() {
    console.log("[WPS] OnApplyForm called");
    // 获取当前文档并应用表单数据
    try {
        const formData = JSON.parse(localStorage.getItem("pevc_wps:formState") || "{}");
        if (Object.keys(formData).length === 0) {
            showWpsMessage("表单数据为空，请先填写表单");
            return true;
        }
        
        applyFormDataQuick(formData);
    } catch (err) {
        console.error("[WPS] Apply form error:", err);
        showWpsMessage("应用失败: " + err.message);
    }
    return true;
}

// =====================================================================
// 快速应用表单数据（不通过 taskpane）
// =====================================================================
async function applyFormDataQuick(formData) {
    const app = wps.WpsApplication();
    const doc = app.ActiveDocument;
    
    if (!doc) {
        showWpsMessage("请先打开一个文档");
        return;
    }
    
    let successCount = 0;
    
    for (const [tag, value] of Object.entries(formData)) {
        if (!value) continue;
        
        // 方法1: 通过 ContentControl
        const ccs = doc.ContentControls;
        for (let i = 1; i <= ccs.Count; i++) {
            const cc = ccs.Item(i);
            if (cc.Tag === tag) {
                cc.Range.Text = String(value);
                successCount++;
                break;
            }
        }
        
        // 方法2: 查找替换
        const patterns = [`【${tag}】`, `[${tag}]`];
        for (const pattern of patterns) {
            const find = doc.Content.Find;
            find.ClearFormatting();
            find.Text = pattern;
            find.Replacement.ClearFormatting();
            find.Replacement.Text = String(value);
            if (find.Execute(pattern, false, false, false, false, false, true, 1, false, String(value), 2)) {
                successCount++;
                break;
            }
        }
    }
    
    showWpsMessage(`已应用 ${successCount} 个字段`);
}

// =====================================================================
// WPS 消息提示
// =====================================================================
function showWpsMessage(msg) {
    try {
        const app = wps.WpsApplication();
        app.Selection.TypeText(""); // 确保有焦点
        console.log("[WPS Message]", msg);
        // WPS 不支持 alert，用 console 输出
    } catch (e) {
        console.log("[WPS Message]", msg);
    }
}

// =====================================================================
// API 能力测试
// =====================================================================
async function runApiTests() {
    console.log("[WPS API Test] Starting tests...");
    
    const app = wps.WpsApplication();
    const doc = app.ActiveDocument;
    
    if (!doc) {
        console.error("[WPS API Test] No active document");
        return false;
    }
    
    let results = [];
    
    // Test 1: Bookmarks
    try {
        doc.Range(0, 0).Text = "书签测试: ";
        const range = doc.Range(0, "书签测试: ".length);
        doc.Bookmarks.Add("TestBookmark", range);
        const bookmark = doc.Bookmarks.Item("TestBookmark");
        if (bookmark) {
            results.push("✅ 书签: 支持");
            bookmark.Delete();
        } else {
            results.push("❌ 书签: 失败");
        }
    } catch (e) {
        results.push("❌ 书签: " + e.message);
    }
    
    // Test 2: Font.Hidden
    try {
        doc.Range(0, 0).InsertParagraphBefore();
        const testText = "隐藏测试文本";
        doc.Range(0, 0).Text = testText;
        const testRange = doc.Range(0, testText.length);
        testRange.Font.Hidden = true;
        testRange.Font.Hidden = false;
        results.push("✅ Font.Hidden: 支持");
    } catch (e) {
        results.push("❌ Font.Hidden: " + e.message);
    }
    
    // Test 3: CustomDocumentProperties
    try {
        const props = doc.CustomDocumentProperties;
        const testName = "PEVC_Test_" + Date.now();
        const testValue = "TestValue";
        props.Add(testName, false, 4, testValue);
        
        let foundProp = null;
        for (let i = 1; i <= props.Count; i++) {
            if (props.Item(i).Name === testName) {
                foundProp = props.Item(i);
                break;
            }
        }
        
        if (foundProp && foundProp.Value === testValue) {
            results.push("✅ CustomDocumentProperties: 支持");
            foundProp.Delete();
        } else {
            results.push("❌ CustomDocumentProperties: 无法读取");
        }
    } catch (e) {
        results.push("❌ CustomDocumentProperties: " + e.message);
    }
    
    // Test 4: ContentControls
    try {
        const sel = app.Selection;
        sel.Text = "内容控件测试";
        const cc = sel.Range.ContentControls.Add(1);
        cc.Tag = "TestCC";
        cc.Title = "测试控件";
        
        const foundCCs = doc.ContentControls;
        let foundTest = false;
        for (let i = 1; i <= foundCCs.Count; i++) {
            if (foundCCs.Item(i).Tag === "TestCC") {
                foundTest = true;
                foundCCs.Item(i).Delete(false);
                break;
            }
        }
        
        if (foundTest) {
            results.push("✅ ContentControls: 支持");
        } else {
            results.push("❌ ContentControls: 无法查找");
        }
    } catch (e) {
        results.push("❌ ContentControls: " + e.message);
    }
    
    // Test 5: Variables
    try {
        const vars = doc.Variables;
        const varName = "PEVC_Test_Var";
        const varValue = "TestVarValue";
        vars.Add(varName, varValue);
        
        const foundVar = vars.Item(varName);
        if (foundVar && foundVar.Value === varValue) {
            results.push("✅ Variables: 支持");
            // Variables 可能没有 Delete 方法，设为空字符串
            foundVar.Value = "";
        } else {
            results.push("❌ Variables: 无法读取");
        }
    } catch (e) {
        results.push("❌ Variables: " + e.message);
    }
    
    // 输出结果
    console.log("\n========== WPS API 测试结果 ==========");
    results.forEach(r => console.log(r));
    console.log("======================================\n");
    
    // 清理测试内容
    try {
        doc.Undo(10);
    } catch (e) {}
    
    return results.every(r => r.startsWith("✅"));
}

// 兼容旧配置的别名
window.OnShowTaskPane = OnShowDialog;
window.OnAddinLoad = OnAddinLoad;
window.OnShowDialog = OnShowDialog;
window.OnApiTest = OnApiTest;
window.OnApplyForm = OnApplyForm;

console.log("[WPS] main.js loaded");
