// Service worker: holds the loopback WebSocket to the MCP server and routes each request
// to a content script running in the right Copilot tab.

const BRIDGE_URL = "ws://127.0.0.1:8765";

const SURFACES = {
  work: { host: "m365.cloud.microsoft", url: "https://m365.cloud.microsoft/chat" },
  m365: { host: "copilot.cloud.microsoft", url: "https://copilot.cloud.microsoft/" },
  web:  { host: "copilot.microsoft.com", url: "https://copilot.microsoft.com/" },
};

let ws = null;
let reconnectDelay = 1000;

function ensureConnected() {
  if (!ws || ws.readyState === WebSocket.CLOSED || ws.readyState === WebSocket.CLOSING) connect();
}

function connect() {
  if (ws && (ws.readyState === WebSocket.OPEN || ws.readyState === WebSocket.CONNECTING)) return;
  ws = new WebSocket(BRIDGE_URL);
  ws.onopen = () => { reconnectDelay = 1000; console.log("[bridge] connected"); };
  ws.onclose = () => {
    console.log("[bridge] closed, retrying in", reconnectDelay, "ms");
    setTimeout(connect, reconnectDelay);
    reconnectDelay = Math.min(reconnectDelay * 2, 30000);
  };
  ws.onerror = () => { try { ws.close(); } catch {} };
  ws.onmessage = async (ev) => {
    let req;
    try { req = JSON.parse(ev.data); } catch { return; }
    try {
      const result = await handle(req);
      send({ id: req.id, result });
    } catch (e) {
      send({ id: req.id, error: String(e && e.message ? e.message : e) });
    }
  };
}

function send(obj) {
  if (ws && ws.readyState === WebSocket.OPEN) ws.send(JSON.stringify(obj));
}

async function handle(req) {
  if (req.type === "ask") return ask(req);
  if (req.type === "surfaces") return surfaces();
  throw new Error("unknown request type: " + req.type);
}

// Find an existing tab for the surface, or open one and wait for it to load.
async function getSurfaceTab(surfaceKey) {
  const s = SURFACES[surfaceKey];
  if (!s) throw new Error("unknown surface: " + surfaceKey);
  const tabs = await chrome.tabs.query({ url: `https://${s.host}/*` });
  if (tabs.length) return tabs[0];
  const tab = await chrome.tabs.create({ url: s.url, active: false });
  await waitForComplete(tab.id);
  // give the SPA a moment to hydrate
  await sleep(2500);
  return tab;
}

function waitForComplete(tabId) {
  return new Promise((resolve) => {
    const listener = (id, info) => {
      if (id === tabId && info.status === "complete") {
        chrome.tabs.onUpdated.removeListener(listener);
        resolve();
      }
    };
    chrome.tabs.onUpdated.addListener(listener);
  });
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function ask(req) {
  const tab = await getSurfaceTab(req.surface);
  return sendToTab(tab.id, {
    cmd: "ask",
    surface: req.surface,
    prompt: req.prompt,
    webGrounding: req.webGrounding,
    newChat: req.newChat,
  });
}

async function surfaces() {
  const out = {};
  for (const [key, s] of Object.entries(SURFACES)) {
    const tabs = await chrome.tabs.query({ url: `https://${s.host}/*` });
    if (!tabs.length) { out[key] = { open: false, host: s.host }; continue; }
    try {
      const info = await sendToTab(tabs[0].id, { cmd: "probe" });
      out[key] = { open: true, host: s.host, ...info };
    } catch (e) {
      out[key] = { open: true, host: s.host, error: String(e.message || e) };
    }
  }
  return out;
}

// Send a message to the content script; inject it first if the tab loaded before the
// extension did.
async function sendToTab(tabId, msg) {
  try {
    return await chrome.tabs.sendMessage(tabId, msg);
  } catch {
    await chrome.scripting.executeScript({ target: { tabId }, files: ["content.js"] });
    await sleep(300);
    return await chrome.tabs.sendMessage(tabId, msg);
  }
}

// MV3 service workers get evicted when idle; a plain setTimeout reconnect can't fire in a dead
// worker. An alarm wakes the worker on a fixed cadence so the bridge always recovers after a
// server restart or worker eviction.
chrome.runtime.onStartup.addListener(ensureConnected);
chrome.runtime.onInstalled.addListener(ensureConnected);
chrome.alarms.create("keepalive", { periodInMinutes: 0.5 });
chrome.alarms.onAlarm.addListener((a) => { if (a.name === "keepalive") ensureConnected(); });

connect();
