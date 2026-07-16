"use strict";

const DEFAULTS = {
  enabled: true,
  autoDefuse: true,
  keepMessages: 2,
  newestPosition: "auto",
  trimQuotes: true,
  quoteCharThreshold: 6000,
  threadNodeThreshold: 3500,
  showToast: true
};

const ids = [
  "enabled",
  "autoDefuse",
  "keepMessages",
  "newestPosition",
  "trimQuotes",
  "quoteCharThreshold",
  "threadNodeThreshold",
  "showToast"
];

const elements = Object.fromEntries(ids.map((id) => [id, document.getElementById(id)]));
const stateBadge = document.getElementById("stateBadge");
const message = document.getElementById("message");
const runButton = document.getElementById("runButton");
const restoreButton = document.getElementById("restoreButton");

function setMessage(text, isError = false) {
  message.textContent = text;
  message.style.color = isError ? "#c4314b" : "";
}

function setBadge(text, state = "") {
  stateBadge.textContent = text;
  stateBadge.className = `badge ${state}`.trim();
}

async function activeTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  return tab;
}

function isOutlookUrl(url = "") {
  try {
    const host = new URL(url).hostname;
    return ["outlook.office.com", "outlook.office365.com", "outlook.live.com", "outlook.cloud.microsoft"].includes(host);
  } catch {
    return false;
  }
}

async function sendToPage(type) {
  const tab = await activeTab();
  if (!tab?.id || !isOutlookUrl(tab.url)) {
    throw new Error("Web版Outlookのタブで実行してください");
  }
  return chrome.tabs.sendMessage(tab.id, { type });
}

function populateSettings(settings) {
  for (const id of ids) {
    const el = elements[id];
    const value = settings[id];
    if (el.type === "checkbox") el.checked = Boolean(value);
    else el.value = String(value);
  }
}

function readValue(element) {
  if (element.type === "checkbox") return element.checked;
  if (element.type === "number") return Number.parseInt(element.value, 10);
  return element.value;
}

async function saveSetting(event) {
  const element = event.currentTarget;
  const value = readValue(element);
  await chrome.storage.local.set({ [element.id]: value });
  setMessage("設定を保存しました");
  window.setTimeout(() => setMessage(""), 1200);
}

function updateStatus(status) {
  document.getElementById("bodyCount").textContent = status.bodyCount?.toLocaleString?.() ?? "-";
  document.getElementById("parkedCount").textContent = status.parked?.toLocaleString?.() ?? "-";
  document.getElementById("removedCount").textContent =
    status.metrics?.nodesRemoved?.toLocaleString?.() ?? "-";
  document.getElementById("pageNodes").textContent = status.pageNodes?.toLocaleString?.() ?? "-";
}

async function refresh() {
  const stored = await chrome.storage.local.get(DEFAULTS);
  populateSettings({ ...DEFAULTS, ...stored });

  const tab = await activeTab();
  if (!tab || !isOutlookUrl(tab.url)) {
    setBadge("対象外", "error");
    runButton.disabled = true;
    restoreButton.disabled = true;
    setMessage("Web版Outlookを開いてから実行してください");
    return;
  }

  try {
    const status = await sendToPage("OTD_STATUS");
    updateStatus(status);
    setBadge("動作中", "ok");
  } catch (error) {
    setBadge("再読込が必要", "error");
    setMessage("拡張機能を読み込んだ後、Outlookタブを再読み込みしてください", true);
  }
}

for (const element of Object.values(elements)) {
  element.addEventListener("change", saveSetting);
}

runButton.addEventListener("click", async () => {
  try {
    setMessage("軽量化しています…");
    const result = await sendToPage("OTD_RUN");
    updateStatus({
      bodyCount: result.metrics?.bodiesFound,
      parked: result.parked,
      metrics: result.metrics
    });
    setMessage("軽量化を実行しました");
  } catch (error) {
    setMessage(error.message || "実行できませんでした", true);
  }
});

restoreButton.addEventListener("click", async () => {
  try {
    const result = await sendToPage("OTD_RESTORE_ALL");
    setMessage(`${result.restored ?? 0}か所を復元しました`);
    await refresh();
  } catch (error) {
    setMessage(error.message || "復元できませんでした", true);
  }
});

refresh().catch((error) => {
  setBadge("エラー", "error");
  setMessage(error.message || "初期化できませんでした", true);
});
