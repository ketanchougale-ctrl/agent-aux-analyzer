const APP_URL = "http://localhost:8501";

const pill       = document.getElementById("pill");
const pillText   = document.getElementById("pill-text");
const btnOpen    = document.getElementById("btnOpen");
const alertBox   = document.getElementById("alertOffline");

function setOnline() {
  pill.className = "pill pill-online";
  pillText.textContent = "Online";
  pill.querySelector(".dot").classList.remove("pulse");
  alertBox.style.display = "none";
  btnOpen.disabled = false;
}

function setOffline() {
  pill.className = "pill pill-offline";
  pillText.textContent = "Offline";
  pill.querySelector(".dot").classList.remove("pulse");
  alertBox.style.display = "block";
  btnOpen.disabled = false; // still allow opening (user may want to try)
}

async function checkServer() {
  try {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), 3000);
    const res = await fetch(APP_URL, {
      method: "GET",
      mode: "no-cors",
      signal: ctrl.signal,
    });
    clearTimeout(timer);
    setOnline();
  } catch {
    setOffline();
  }
}

btnOpen.addEventListener("click", async () => {
  const tabs = await chrome.tabs.query({ url: `${APP_URL}/*` });
  if (tabs.length > 0) {
    await chrome.tabs.update(tabs[0].id, { active: true });
    const win = await chrome.windows.get(tabs[0].windowId);
    chrome.windows.update(win.id, { focused: true });
  } else {
    chrome.tabs.create({ url: APP_URL });
  }
  window.close();
});

checkServer();
