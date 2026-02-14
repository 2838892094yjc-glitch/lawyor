/**
 * Runtime config loader for cloud-first deployment.
 * Supports one-time install with remote manifest + API endpoint sync.
 */
(function (global) {
  "use strict";

  var STORAGE_KEY = "pevc_runtime_config_v1";
  var DEFAULTS = {
    apiBaseUrl: "http://127.0.0.1:8765",
    manifestUrl: "",
    taskpaneUrl: "",
    autoRefreshMs: 5 * 60 * 1000,
    lastSyncAt: 0,
    version: "local"
  };

  function normalizeBaseUrl(url) {
    if (!url || typeof url !== "string") return "";
    return url.replace(/\/+$/, "");
  }

  function readLocalConfig() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return null;
      var parsed = JSON.parse(raw);
      return parsed && typeof parsed === "object" ? parsed : null;
    } catch (_err) {
      return null;
    }
  }

  function writeLocalConfig(config) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(config));
      return true;
    } catch (_err) {
      return false;
    }
  }

  function mergeConfig(base, extra) {
    var merged = Object.assign({}, base || {}, extra || {});
    merged.apiBaseUrl = normalizeBaseUrl(merged.apiBaseUrl || DEFAULTS.apiBaseUrl);
    merged.manifestUrl = merged.manifestUrl || "";
    merged.taskpaneUrl = merged.taskpaneUrl || "";
    if (!merged.autoRefreshMs || merged.autoRefreshMs < 30000) {
      merged.autoRefreshMs = DEFAULTS.autoRefreshMs;
    }
    return merged;
  }

  function getQueryParam(name) {
    try {
      if (!global.location || !global.location.search) return "";
      var params = new URLSearchParams(global.location.search);
      return (params.get(name) || "").trim();
    } catch (_err) {
      return "";
    }
  }

  function readBootstrapParamsFromUrl() {
    var manifestUrl = getQueryParam("pevc_manifest");
    var apiBaseUrl = getQueryParam("pevc_api_base");
    var taskpaneUrl = getQueryParam("pevc_taskpane");
    var patch = {};

    if (manifestUrl) patch.manifestUrl = manifestUrl;
    if (apiBaseUrl) patch.apiBaseUrl = apiBaseUrl;
    if (taskpaneUrl) patch.taskpaneUrl = taskpaneUrl;

    return patch;
  }

  var runtimeConfig = mergeConfig(DEFAULTS, readLocalConfig());
  var bootstrapPatch = readBootstrapParamsFromUrl();
  if (bootstrapPatch.manifestUrl || bootstrapPatch.apiBaseUrl || bootstrapPatch.taskpaneUrl) {
    runtimeConfig = mergeConfig(runtimeConfig, bootstrapPatch);
    writeLocalConfig(runtimeConfig);
  }

  function getConfig() {
    return Object.assign({}, runtimeConfig);
  }

  function setConfig(next, options) {
    runtimeConfig = mergeConfig(runtimeConfig, next || {});
    if (!options || options.persist !== false) {
      writeLocalConfig(runtimeConfig);
    }
    return getConfig();
  }

  function getApiBaseUrl() {
    return normalizeBaseUrl(runtimeConfig.apiBaseUrl || DEFAULTS.apiBaseUrl);
  }

  function buildApiUrl(path) {
    var base = getApiBaseUrl();
    var cleanPath = path && path.charAt(0) === "/" ? path : "/" + (path || "");
    return base + cleanPath;
  }

  async function loadRemoteConfig(options) {
    var opts = options || {};
    var manifestUrl = opts.manifestUrl || runtimeConfig.manifestUrl;
    if (!manifestUrl) {
      return { ok: false, reason: "manifest_url_empty", config: getConfig() };
    }
    if (typeof fetch !== "function") {
      return { ok: false, reason: "fetch_unavailable", config: getConfig() };
    }

    var timeoutMs = typeof opts.timeoutMs === "number" ? opts.timeoutMs : 3500;
    var controller = typeof AbortController === "function" ? new AbortController() : null;
    var timer = setTimeout(function () {
      if (controller) controller.abort();
    }, timeoutMs);

    try {
      var response = await fetch(manifestUrl, {
        cache: "no-store",
        signal: controller ? controller.signal : undefined
      });

      if (!response.ok) {
        throw new Error("manifest http " + response.status);
      }

      var payload = await response.json();
      var remotePatch = {
        version: payload.version || payload.latestVersion || runtimeConfig.version,
        taskpaneUrl: payload.taskpaneUrl || (payload.client && payload.client.taskpaneUrl) || runtimeConfig.taskpaneUrl,
        apiBaseUrl: payload.apiBaseUrl || (payload.services && payload.services.apiBaseUrl) || (payload.services && payload.services.aiBaseUrl) || runtimeConfig.apiBaseUrl,
        autoRefreshMs: payload.autoRefreshMs || (payload.client && payload.client.autoRefreshMs) || runtimeConfig.autoRefreshMs,
        lastSyncAt: Date.now()
      };

      setConfig(remotePatch, { persist: true });
      return { ok: true, config: getConfig(), payload: payload };
    } catch (err) {
      return {
        ok: false,
        reason: err && err.name === "AbortError" ? "timeout" : "fetch_failed",
        error: err,
        config: getConfig()
      };
    } finally {
      clearTimeout(timer);
    }
  }

  var autoSyncTimer = null;
  function startAutoSync() {
    if (!runtimeConfig.manifestUrl) return;
    if (autoSyncTimer) clearInterval(autoSyncTimer);
    autoSyncTimer = setInterval(function () {
      loadRemoteConfig({ timeoutMs: 2500 });
    }, runtimeConfig.autoRefreshMs || DEFAULTS.autoRefreshMs);
  }

  global.PEVCRuntimeConfig = {
    STORAGE_KEY: STORAGE_KEY,
    getConfig: getConfig,
    setConfig: setConfig,
    getApiBaseUrl: getApiBaseUrl,
    buildApiUrl: buildApiUrl,
    loadRemoteConfig: loadRemoteConfig,
    startAutoSync: startAutoSync
  };
})(typeof window !== "undefined" ? window : globalThis);
