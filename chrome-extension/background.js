const APP_URL      = "http://localhost:8501";
const LAUNCH_PROTO = "streamlit-aux://start";
const POLL_MS      = 2000;
const MAX_WAIT_MS  = 90000;

async function serverOnline() {
  try {
    const ctrl = new AbortController();
    const id   = setTimeout(() => ctrl.abort(), 3000);
    await fetch(APP_URL, { method: "GET", mode: "no-cors", signal: ctrl.signal });
    clearTimeout(id);
    return true;
  } catch {
    return false;
  }
}

async function openApp() {
  const tabs = await chrome.tabs.query({ url: `${APP_URL}/*` });
  if (tabs.length > 0) {
    await chrome.tabs.update(tabs[0].id, { active: true });
    chrome.windows.update(tabs[0].windowId, { focused: true });
  } else {
    chrome.tabs.create({ url: APP_URL });
  }
}

async function startAndOpen() {
  if (await serverOnline()) {
    openApp();
    return;
  }

  let protocolTabId = null;
  try {
    const t = await chrome.tabs.create({ url: LAUNCH_PROTO, active: true });
    protocolTabId = t.id;
  } catch (e) {
    console.error("[AUX] Failed to open launch protocol:", e);
  }

  const deadline = Date.now() + MAX_WAIT_MS;
  while (Date.now() < deadline) {
    await new Promise(r => setTimeout(r, POLL_MS));
    if (await serverOnline()) {
      if (protocolTabId !== null) {
        chrome.tabs.remove(protocolTabId).catch(() => {});
      }
      openApp();
      return;
    }
  }

  if (protocolTabId !== null) {
    chrome.tabs.remove(protocolTabId).catch(() => {});
  }
}

chrome.runtime.onMessage.addListener((msg) => {
  if (msg.action === "startAndOpen") startAndOpen();
  if (msg.action === "openApp")      openApp();
});
