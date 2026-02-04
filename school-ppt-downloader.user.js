// ==UserScript==
// @name         School PPT Downloader
// @namespace    local.school.ppt
// @version      0.1.0
// @description  Extract and download PPT files from course platform pages
// @match        *://*/*
// @grant        GM_download
// @grant        GM_setClipboard
// @grant        GM_registerMenuCommand
// @grant        GM_getValue
// @grant        GM_setValue
// @run-at       document-idle
// ==/UserScript==

(() => {
  "use strict";

  if (window.top !== window.self) return;

  const state = {
    items: new Map(),
    lastScanAt: 0,
    deepModeEnabled: false,
    observer: null,
  };

  const storageKeys = {
    whitelist: "pptWhitelist",
    mode: "pptMode",
    deepMode: "pptDeepMode",
  };

  const extPattern = /\.(pptx?|ppsx?|potx?|pdf)(\?|#|$)/i;
  const urlPattern = /https?:\/\/[^\s"'<>]+?\.(pptx?|ppsx?|potx?|pdf)(\?[^\s"'<>]*)?/gi;

  function normalizeUrl(raw) {
    try {
      return new URL(raw, location.href).href;
    } catch {
      return null;
    }
  }

  function getFileName(url) {
    try {
      const clean = url.split(/[?#]/)[0];
      const name = decodeURIComponent(clean.substring(clean.lastIndexOf("/") + 1));
      return name || "slide.ppt";
    } catch {
      return "slide.ppt";
    }
  }

  function addItem(url, source) {
    if (!url || state.items.has(url)) return;
    state.items.set(url, { url, source, addedAt: Date.now() });
    renderList();
  }

  function considerUrl(raw, source) {
    if (!raw || !extPattern.test(raw)) return;
    const url = normalizeUrl(raw);
    if (!url) return;
    addItem(url, source);
  }

  function decodeOnlinePreview(raw) {
    try {
      const url = new URL(raw, location.href);
      if (!/\/onlinePreview/i.test(url.pathname)) return null;
      const encoded = url.searchParams.get("url");
      if (!encoded) return null;
      const decoded = atob(encoded);
      return normalizeUrl(decoded);
    } catch {
      return null;
    }
  }

  function extractFromText(text, source) {
    if (!text) return;
    const matches = text.match(urlPattern) || [];
    matches.forEach((m) => addItem(normalizeUrl(m), source));
  }

  function scanDom() {
    const now = Date.now();
    if (now - state.lastScanAt < 800) return;
    state.lastScanAt = now;

    const elements = document.querySelectorAll("a[href], iframe[src], embed[src], object[data]");
    elements.forEach((el) => {
      const url = el.getAttribute("href") || el.getAttribute("src") || el.getAttribute("data");
      const decoded = decodeOnlinePreview(url);
      if (decoded) {
        addItem(decoded, "preview");
        return;
      }
      considerUrl(url, "dom");
    });

    const textNodes = document.querySelectorAll("script, pre");
    textNodes.forEach((el) => extractFromText(el.textContent, "inline"));
  }

  function scanResources() {
    try {
      const entries = performance.getEntriesByType("resource");
      entries.forEach((entry) => considerUrl(entry.name, "resource"));
    } catch {
      // ignore
    }
  }

  function hookFetch() {
    if (!window.fetch) return;
    const originalFetch = window.fetch;
    window.fetch = async (...args) => {
      const response = await originalFetch(...args);
      try {
        const clone = response.clone();
        const contentType = clone.headers.get("content-type") || "";
        if (contentType.includes("application/json") || contentType.includes("text")) {
          const text = await clone.text();
          extractFromText(text, "fetch");
        }
      } catch {
        // ignore
      }
      return response;
    };
  }

  function hookXhr() {
    const open = XMLHttpRequest.prototype.open;
    const send = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.open = function (...args) {
      this._pptUrl = args[1];
      return open.apply(this, args);
    };
    XMLHttpRequest.prototype.send = function (...args) {
      this.addEventListener("load", () => {
        try {
          if (typeof this.responseText === "string") {
            extractFromText(this.responseText, "xhr");
          }
        } catch {
          // ignore
        }
      });
      return send.apply(this, args);
    };
  }

  function downloadUrl(url) {
    const name = getFileName(url);
    try {
      GM_download({ url, name, saveAs: true, onerror: () => window.open(url, "_blank") });
    } catch {
      window.open(url, "_blank");
    }
  }

  function copyAll() {
    const lines = Array.from(state.items.values()).map((i) => i.url).join("\n");
    GM_setClipboard(lines || "");
    toast(lines ? "已复制全部链接" : "暂无链接");
  }

  function createPanel() {
    const panel = document.createElement("div");
    panel.id = "ppt-downloader-panel";
    panel.classList.add("hidden");
    panel.innerHTML = `
      <div class="ppt-header">
        <div class="ppt-title">PPT 下载</div>
        <div class="ppt-actions">
          <button data-action="scan">扫描</button>
          <button data-action="copy">复制</button>
          <button data-action="close">收起</button>
        </div>
      </div>
      <div class="ppt-body">
        <div class="ppt-empty">尚未发现PPT链接</div>
        <ul class="ppt-list"></ul>
      </div>
    `;
    document.body.appendChild(panel);

    panel.addEventListener("click", (event) => {
      const action = event.target?.getAttribute("data-action");
      if (!action) return;
      if (action === "scan") scanOnce();
      if (action === "copy") copyAll();
      if (action === "close") panel.classList.add("hidden");
    });

    return panel;
  }

  function createButton() {
    const btn = document.createElement("button");
    btn.id = "ppt-downloader-btn";
    btn.textContent = "P";
    btn.addEventListener("click", () => {
      const panel = document.getElementById("ppt-downloader-panel");
      panel?.classList.toggle("hidden");
    });
    document.body.appendChild(btn);
  }

  function renderList() {
    const panel = document.getElementById("ppt-downloader-panel");
    if (!panel) return;
    const list = panel.querySelector(".ppt-list");
    const empty = panel.querySelector(".ppt-empty");
    if (!list || !empty) return;

    list.innerHTML = "";
    const items = Array.from(state.items.values());
    empty.style.display = items.length ? "none" : "block";
    items.forEach((item, index) => {
      const li = document.createElement("li");
      const name = getFileName(item.url);
      li.innerHTML = `
        <span class="ppt-index">${index + 1}</span>
        <span class="ppt-name" title="${name}">${name}</span>
        <button data-url="${item.url}">下载</button>
      `;
      li.querySelector("button").addEventListener("click", () => downloadUrl(item.url));
      list.appendChild(li);
    });
  }

  function toast(text) {
    let el = document.getElementById("ppt-downloader-toast");
    if (!el) {
      el = document.createElement("div");
      el.id = "ppt-downloader-toast";
      document.body.appendChild(el);
    }
    el.textContent = text;
    el.classList.add("show");
    clearTimeout(el._timer);
    el._timer = setTimeout(() => el.classList.remove("show"), 1800);
  }

  function addStyles() {
    const style = document.createElement("style");
    style.textContent = `
      #ppt-downloader-btn {
        position: fixed;
        right: 16px;
        bottom: 18px;
        z-index: 99999;
        width: 44px;
        height: 44px;
        border: 0;
        border-radius: 999px;
        background: #1f6feb;
        color: #fff;
        font-size: 16px;
        font-weight: 600;
        cursor: pointer;
        box-shadow: 0 10px 24px rgba(31, 111, 235, 0.35);
      }
      #ppt-downloader-panel {
        position: fixed;
        right: 16px;
        bottom: 68px;
        width: 360px;
        max-height: 50vh;
        background: #fff;
        border-radius: 12px;
        box-shadow: 0 20px 40px rgba(0, 0, 0, 0.2);
        z-index: 99999;
        display: flex;
        flex-direction: column;
        overflow: hidden;
        font-family: "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
      }
      #ppt-downloader-panel.hidden {
        display: none;
      }
      #ppt-downloader-panel.collapsed {
        max-height: 42px;
      }
      #ppt-downloader-panel .ppt-header {
        padding: 10px 12px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        background: #f6f8fa;
      }
      #ppt-downloader-panel .ppt-title {
        font-weight: 600;
        font-size: 14px;
        color: #1f2328;
      }
      #ppt-downloader-panel .ppt-actions button {
        margin-left: 6px;
        padding: 6px 10px;
        border-radius: 8px;
        border: 1px solid #d0d7de;
        background: #fff;
        cursor: pointer;
        font-size: 12px;
      }
      #ppt-downloader-panel .ppt-body {
        padding: 10px 12px;
        overflow: auto;
      }
      #ppt-downloader-panel .ppt-empty {
        color: #6a737d;
        font-size: 13px;
        padding: 8px 0;
      }
      #ppt-downloader-panel .ppt-list {
        list-style: none;
        margin: 0;
        padding: 0;
      }
      #ppt-downloader-panel .ppt-list li {
        display: grid;
        grid-template-columns: 20px 1fr 64px;
        gap: 8px;
        align-items: center;
        padding: 6px 0;
        border-bottom: 1px dashed #e6e8eb;
      }
      #ppt-downloader-panel .ppt-index {
        color: #8c959f;
        font-size: 12px;
      }
      #ppt-downloader-panel .ppt-name {
        font-size: 13px;
        color: #1f2328;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }
      #ppt-downloader-panel .ppt-list button {
        padding: 6px 8px;
        border-radius: 6px;
        border: 1px solid #1f6feb;
        background: #1f6feb;
        color: #fff;
        cursor: pointer;
        font-size: 12px;
      }
      #ppt-downloader-toast {
        position: fixed;
        right: 20px;
        bottom: 140px;
        padding: 8px 12px;
        background: rgba(32, 33, 36, 0.92);
        color: #fff;
        border-radius: 8px;
        font-size: 12px;
        opacity: 0;
        transform: translateY(6px);
        transition: all 0.18s ease;
        z-index: 99999;
        pointer-events: none;
      }
      #ppt-downloader-toast.show {
        opacity: 1;
        transform: translateY(0);
      }
    `;
    document.head.appendChild(style);
  }

  function scanOnce() {
    scanDom();
    scanResources();
  }

  function isDeepMode() {
    return GM_getValue(storageKeys.deepMode, false);
  }

  function setDeepMode(next) {
    GM_setValue(storageKeys.deepMode, next);
  }

  function enableDeepMode() {
    if (state.deepModeEnabled) return;
    state.deepModeEnabled = true;
    hookFetch();
    hookXhr();
    scanOnce();
    state.observer = new MutationObserver(() => scanDom());
    state.observer.observe(document.documentElement, { childList: true, subtree: true });
    toast("深度抓取已启用，建议刷新页面");
  }

  function disableDeepMode() {
    setDeepMode(false);
    if (state.observer) {
      state.observer.disconnect();
      state.observer = null;
    }
    state.deepModeEnabled = false;
    toast("深度抓取已关闭");
  }

  function toggleDeepMode() {
    const next = !isDeepMode();
    setDeepMode(next);
    if (!isEnabledForHost()) {
      toast(next ? "深度抓取已开启(此站点未启用)" : "深度抓取已关闭");
      return;
    }
    if (next) {
      enableDeepMode();
      return;
    }
    disableDeepMode();
  }

  function getHost() {
    return location.hostname;
  }

  function getMode() {
    return GM_getValue(storageKeys.mode, "all");
  }

  function setMode(mode) {
    GM_setValue(storageKeys.mode, mode);
  }

  function getWhitelist() {
    return GM_getValue(storageKeys.whitelist, []);
  }

  function setWhitelist(list) {
    GM_setValue(storageKeys.whitelist, list);
  }

  function isEnabledForHost() {
    const mode = getMode();
    if (mode !== "whitelist") return true;
    const list = getWhitelist();
    return list.includes(getHost());
  }

  function enableOnlyThisHost() {
    const host = getHost();
    const list = getWhitelist();
    if (!list.includes(host)) {
      list.push(host);
      setWhitelist(list);
    }
    setMode("whitelist");
    toast(`已仅在 ${host} 启用`);
  }

  function disableThisHost() {
    const host = getHost();
    const list = getWhitelist().filter((item) => item !== host);
    setWhitelist(list);
    if (list.length === 0) setMode("all");
    toast(`已移除 ${host}`);
  }

  function enableAll() {
    setMode("all");
    toast("已对所有网站启用");
  }

  function showStatus() {
    const mode = getMode();
    const deep = isDeepMode() ? "深度抓取:开" : "深度抓取:关";
    if (mode === "all") {
      toast(`当前模式：全站启用，${deep}`);
      return;
    }
    const list = getWhitelist();
    const host = getHost();
    const current = list.includes(host) ? "(当前站点已启用)" : "(当前站点未启用)";
    toast(`当前模式：仅白名单 ${current}，${deep}`);
  }

  function showWhitelist() {
    const list = getWhitelist();
    if (!list.length) {
      toast("白名单为空");
      return;
    }
    const text = `白名单(${list.length})\n\n${list.join("\n")}`;
    window.alert(text);
  }

  function clearWhitelist() {
    if (!window.confirm("确认清空白名单？")) return;
    setWhitelist([]);
    setMode("all");
    toast("已清空白名单，恢复全站启用");
  }

  function init() {
    if (!isEnabledForHost()) {
      GM_registerMenuCommand("查看当前模式", showStatus);
      GM_registerMenuCommand("查看白名单", showWhitelist);
      GM_registerMenuCommand("清空白名单", clearWhitelist);
      GM_registerMenuCommand("仅在此站点启用", enableOnlyThisHost);
      GM_registerMenuCommand("启用所有网站", enableAll);
      GM_registerMenuCommand("切换深度抓取", toggleDeepMode);
      return;
    }

    addStyles();
    createPanel();
    createButton();

    if (isDeepMode()) enableDeepMode();

    GM_registerMenuCommand("复制全部PPT链接", copyAll);
    GM_registerMenuCommand("查看当前模式", showStatus);
    GM_registerMenuCommand("查看白名单", showWhitelist);
    GM_registerMenuCommand("清空白名单", clearWhitelist);
    GM_registerMenuCommand("仅在此站点启用", enableOnlyThisHost);
    GM_registerMenuCommand("移除此站点", disableThisHost);
    GM_registerMenuCommand("启用所有网站", enableAll);
    GM_registerMenuCommand("切换深度抓取", toggleDeepMode);
  }

  init();
})();
