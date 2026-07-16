(() => {
  "use strict";

  const DEFAULTS = Object.freeze({
    enabled: true,
    autoDefuse: true,
    keepMessages: 2,
    newestPosition: "auto",
    trimQuotes: true,
    quoteCharThreshold: 6000,
    threadNodeThreshold: 3500,
    showToast: true,
    debug: false
  });

  const BODY_SELECTORS = [
    '[aria-label="Message body"]',
    '[aria-label="メッセージ本文"]',
    '[data-testid="message-body"]',
    'div[role="document"]'
  ].join(",");

  const DIRECT_QUOTE_SELECTORS = [
    '#divRplyFwdMsg',
    '[id^="divRplyFwdMsg"]',
    'blockquote[type="cite"]',
    '.gmail_quote',
    '.yahoo_quoted'
  ].join(",");

  const HEADER_MARKERS = [
    /(?:^|\n)\s*(?:from|差出人|送信者)\s*[:：]/i,
    /(?:^|\n)\s*(?:sent|date|送信日時|日時)\s*[:：]/i,
    /(?:^|\n)\s*(?:to|宛先)\s*[:：]/i,
    /(?:^|\n)\s*(?:subject|件名)\s*[:：]/i,
    /(?:^|\n)\s*(?:cc|ＣＣ)\s*[:：]/i
  ];

  let settings = { ...DEFAULTS };
  let observer = null;
  let debounceTimer = null;
  let processing = false;
  let runCounter = 0;
  let pauseUntil = 0;

  const parkedSections = new Map();

  const metrics = {
    runs: 0,
    bodiesFound: 0,
    messagesParked: 0,
    quotesParked: 0,
    nodesRemoved: 0,
    charsParked: 0,
    lastRunAt: null,
    lastReason: null
  };

  function log(...args) {
    if (settings.debug) {
      console.debug("[Outlook Thread Defuser]", ...args);
    }
  }

  function normalizeInteger(value, min, max, fallback) {
    const n = Number.parseInt(value, 10);
    if (!Number.isFinite(n)) return fallback;
    return Math.min(max, Math.max(min, n));
  }

  async function loadSettings() {
    const stored = await chrome.storage.local.get(DEFAULTS);
    settings = {
      ...DEFAULTS,
      ...stored,
      keepMessages: normalizeInteger(stored.keepMessages, 1, 8, DEFAULTS.keepMessages),
      quoteCharThreshold: normalizeInteger(
        stored.quoteCharThreshold,
        1000,
        100000,
        DEFAULTS.quoteCharThreshold
      ),
      threadNodeThreshold: normalizeInteger(
        stored.threadNodeThreshold,
        500,
        30000,
        DEFAULTS.threadNodeThreshold
      )
    };
  }

  function countNodes(root) {
    if (!root) return 0;
    let count = 0;
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT);
    while (walker.nextNode()) count += 1;
    return count;
  }

  function isUsableBody(element) {
    if (!(element instanceof HTMLElement)) return false;
    if (!element.isConnected) return false;
    if (element.closest(".otd-placeholder")) return false;
    if (element.closest('[contenteditable="true"]')) return false;
    if (element.matches('[contenteditable="true"]')) return false;

    const textLength = (element.innerText || "").trim().length;
    const nodeCount = element.querySelectorAll("*").length;
    if (textLength < 20 && nodeCount < 10) return false;

    const rect = element.getBoundingClientRect();
    if (rect.width > 0 && rect.width < 180) return false;
    return true;
  }

  function findMessageBodies() {
    const all = [...document.querySelectorAll(BODY_SELECTORS)].filter(isUsableBody);

    const deduped = all.filter((candidate) => {
      return !all.some(
        (other) => other !== candidate && candidate.contains(other) && isUsableBody(other)
      );
    });

    return deduped.sort((a, b) => {
      const ar = a.getBoundingClientRect();
      const br = b.getBoundingClientRect();
      return ar.top - br.top || ar.left - br.left;
    });
  }

  function distanceFromViewport(element) {
    const rect = element.getBoundingClientRect();
    const vh = window.innerHeight || document.documentElement.clientHeight;
    if (rect.bottom >= 0 && rect.top <= vh) return 0;
    if (rect.bottom < 0) return Math.abs(rect.bottom);
    return Math.abs(rect.top - vh);
  }

  function selectBodiesToKeep(bodies) {
    const keepCount = Math.min(settings.keepMessages, bodies.length);
    if (keepCount >= bodies.length) return new Set(bodies);

    if (settings.newestPosition === "top") {
      return new Set(bodies.slice(0, keepCount));
    }
    if (settings.newestPosition === "bottom") {
      return new Set(bodies.slice(-keepCount));
    }

    return new Set(
      [...bodies]
        .sort((a, b) => distanceFromViewport(a) - distanceFromViewport(b))
        .slice(0, keepCount)
    );
  }

  function createPlaceholder({ kind, title, details, html, nodeCount }) {
    const placeholder = document.createElement("section");
    placeholder.className = `otd-placeholder otd-placeholder--${kind}`;
    placeholder.setAttribute("data-otd-placeholder", "true");

    const label = document.createElement("span");
    label.className = "otd-placeholder__label";
    label.textContent = title;

    const meta = document.createElement("span");
    meta.className = "otd-placeholder__meta";
    meta.textContent = details;

    const button = document.createElement("button");
    button.type = "button";
    button.className = "otd-placeholder__button";
    button.textContent = "この部分を復元";

    placeholder.append(label, meta, button);

    const id = `otd-${Date.now()}-${runCounter++}`;
    placeholder.dataset.otdId = id;
    parkedSections.set(id, { html, nodeCount, kind });

    button.addEventListener("click", () => restorePlaceholder(placeholder));
    return placeholder;
  }

  function htmlToFragment(html, contextNode) {
    const range = document.createRange();
    range.selectNode(contextNode || document.body);
    return range.createContextualFragment(html);
  }

  function restorePlaceholder(placeholder, pauseMs = 120000) {
    const id = placeholder.dataset.otdId;
    const parked = parkedSections.get(id);
    if (!parked) return;

    try {
      const fragment = htmlToFragment(parked.html, placeholder);
      placeholder.replaceWith(fragment);
      parkedSections.delete(id);
      pauseUntil = Math.max(pauseUntil, Date.now() + pauseMs);
    } catch (error) {
      console.error("[Outlook Thread Defuser] 復元に失敗しました", error);
    }
  }

  function parkWholeBody(body) {
    if (!body.isConnected || body.querySelector(":scope > .otd-placeholder--message")) return false;

    const html = body.innerHTML;
    if (!html.trim()) return false;

    const nodeCount = body.querySelectorAll("*").length;
    const charCount = (body.innerText || "").length;
    const placeholder = createPlaceholder({
      kind: "message",
      title: "古いメッセージ本文を退避しました",
      details: `${nodeCount.toLocaleString()}ノード / 約${charCount.toLocaleString()}文字`,
      html,
      nodeCount
    });

    body.replaceChildren(placeholder);
    metrics.messagesParked += 1;
    metrics.nodesRemoved += nodeCount;
    metrics.charsParked += charCount;
    return true;
  }

  function markerScore(text) {
    const sample = text.slice(0, 1800);
    return HEADER_MARKERS.reduce((score, regex) => score + (regex.test(sample) ? 1 : 0), 0);
  }

  function findHeaderQuoteStart(body) {
    const bodyTextLength = (body.innerText || "").length;
    if (bodyTextLength < settings.quoteCharThreshold) return null;

    const candidates = [...body.querySelectorAll("div, table, p")];
    for (const candidate of candidates) {
      if (candidate.closest(".otd-placeholder")) continue;
      const text = (candidate.innerText || "").trim();
      if (text.length < 40 || text.length > Math.max(6000, bodyTextLength * 0.95)) continue;
      if (markerScore(text) < 3) continue;

      const range = document.createRange();
      try {
        range.setStart(body, 0);
        range.setEndBefore(candidate);
        const precedingLength = range.toString().trim().length;
        if (precedingLength < 180) continue;
      } catch {
        continue;
      }

      return candidate;
    }
    return null;
  }

  function topLevelBranch(body, node) {
    let current = node;
    while (current && current.parentElement && current.parentElement !== body) {
      current = current.parentElement;
    }
    return current && current.parentElement === body ? current : null;
  }

  function collectFromNodeToEnd(startNode) {
    const nodes = [];
    let current = startNode;
    while (current) {
      nodes.push(current);
      current = current.nextSibling;
    }
    return nodes;
  }

  function parkNodeRange(body, startNode, reason) {
    const branch = topLevelBranch(body, startNode);
    if (!branch) return false;

    const nodes = collectFromNodeToEnd(branch);
    if (nodes.length === 0) return false;

    const wrapper = document.createElement("div");
    for (const node of nodes) {
      wrapper.appendChild(node.cloneNode(true));
    }

    const html = wrapper.innerHTML;
    const text = wrapper.innerText || "";
    const nodeCount = wrapper.querySelectorAll("*").length;
    if (text.trim().length < 200 && nodeCount < 10) return false;

    const placeholder = createPlaceholder({
      kind: "quote",
      title: "返信履歴を退避しました",
      details: `${reason}・${nodeCount.toLocaleString()}ノード / 約${text.length.toLocaleString()}文字`,
      html,
      nodeCount
    });

    branch.before(placeholder);
    for (const node of nodes) node.remove();

    metrics.quotesParked += 1;
    metrics.nodesRemoved += nodeCount;
    metrics.charsParked += text.length;
    return true;
  }

  function trimQuotedHistory(body) {
    if (!settings.trimQuotes || !body.isConnected) return false;
    if (body.querySelector(":scope > .otd-placeholder--quote")) return false;

    const textLength = (body.innerText || "").length;
    if (textLength < settings.quoteCharThreshold) return false;

    const direct = [...body.querySelectorAll(DIRECT_QUOTE_SELECTORS)].find(
      (element) => !element.closest(".otd-placeholder")
    );
    if (direct && parkNodeRange(body, direct, "引用要素を検出")) return true;

    const headerStart = findHeaderQuoteStart(body);
    if (headerStart && parkNodeRange(body, headerStart, "メールヘッダーを検出")) return true;

    return false;
  }

  function shouldDefuse(bodies) {
    if (bodies.length > settings.keepMessages) return true;
    const approximateNodes = bodies.reduce(
      (sum, body) => sum + body.querySelectorAll("*").length,
      0
    );
    return approximateNodes >= settings.threadNodeThreshold;
  }

  function showToast(delta) {
    if (!settings.showToast || delta <= 0) return;
    document.querySelector(".otd-toast")?.remove();

    const toast = document.createElement("div");
    toast.className = "otd-toast";
    toast.textContent = `Thread Defuser: ${delta.toLocaleString()}個のDOMノードを退避`;
    document.body.appendChild(toast);
    requestAnimationFrame(() => toast.classList.add("otd-toast--visible"));
    window.setTimeout(() => {
      toast.classList.remove("otd-toast--visible");
      window.setTimeout(() => toast.remove(), 250);
    }, 2600);
  }

  function defuse(reason = "manual") {
    if (!settings.enabled || processing) return;
    if (Date.now() < pauseUntil && reason !== "popup") return;
    processing = true;

    try {
      const before = metrics.nodesRemoved;
      const bodies = findMessageBodies();
      metrics.runs += 1;
      metrics.bodiesFound = bodies.length;
      metrics.lastRunAt = new Date().toISOString();
      metrics.lastReason = reason;

      if (bodies.length === 0) return;

      for (const body of bodies) {
        trimQuotedHistory(body);
      }

      if (shouldDefuse(bodies)) {
        const keep = selectBodiesToKeep(bodies);
        for (const body of bodies) {
          if (!keep.has(body)) {
            parkWholeBody(body);
          }
        }
      }

      const delta = metrics.nodesRemoved - before;
      showToast(delta);
      log("defused", { reason, bodies: bodies.length, delta, metrics: { ...metrics } });
    } finally {
      processing = false;
    }
  }

  function scheduleDefuse(reason = "mutation", delay = 450) {
    if (!settings.enabled || !settings.autoDefuse) return;
    if (Date.now() < pauseUntil) return;
    window.clearTimeout(debounceTimer);
    debounceTimer = window.setTimeout(() => {
      const run = () => defuse(reason);
      if ("requestIdleCallback" in window) {
        window.requestIdleCallback(run, { timeout: 1200 });
      } else {
        run();
      }
    }, delay);
  }

  function startObserver() {
    observer?.disconnect();
    observer = new MutationObserver((mutations) => {
      if (processing) return;
      const meaningful = mutations.some((mutation) => {
        return [...mutation.addedNodes].some(
          (node) => node instanceof HTMLElement && !node.closest?.(".otd-placeholder")
        );
      });
      if (meaningful) scheduleDefuse("mutation");
    });
    observer.observe(document.documentElement, { childList: true, subtree: true });
  }

  function restoreAll() {
    const placeholders = [...document.querySelectorAll(".otd-placeholder[data-otd-id]")];
    pauseUntil = Date.now() + 300000;
    observer?.disconnect();
    processing = true;
    try {
      for (const placeholder of placeholders) restorePlaceholder(placeholder, 300000);
    } finally {
      processing = false;
      startObserver();
    }
    return placeholders.length;
  }

  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    if (!message || typeof message.type !== "string") return;

    if (message.type === "OTD_RUN") {
      defuse("popup");
      sendResponse({ ok: true, metrics: { ...metrics }, parked: parkedSections.size });
      return;
    }

    if (message.type === "OTD_STATUS") {
      sendResponse({
        ok: true,
        settings: { ...settings },
        metrics: { ...metrics },
        parked: parkedSections.size,
        bodyCount: findMessageBodies().length,
        pageNodes: countNodes(document.body)
      });
      return;
    }

    if (message.type === "OTD_RESTORE_ALL") {
      const restored = restoreAll();
      sendResponse({ ok: true, restored });
      return;
    }
  });

  chrome.storage.onChanged.addListener(async (changes, areaName) => {
    if (areaName !== "local") return;
    for (const [key, change] of Object.entries(changes)) {
      settings[key] = change.newValue;
    }
    if (settings.enabled && settings.autoDefuse) scheduleDefuse("settings", 150);
  });

  async function init() {
    await loadSettings();
    startObserver();
    if (settings.enabled && settings.autoDefuse) scheduleDefuse("initial", 700);
    log("initialized", settings);
  }

  init().catch((error) => {
    console.error("[Outlook Thread Defuser] 初期化に失敗しました", error);
  });
})();
