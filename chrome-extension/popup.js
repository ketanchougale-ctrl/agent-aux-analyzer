const APP_URL  = "http://localhost:8501";

const pill     = document.getElementById("pill");
const pillText = document.getElementById("pill-text");
const btnOpen  = document.getElementById("btnOpen");
const alertBox = document.getElementById("alertOffline");

let _online = false;

function setOnline() {
  pill.className = "pill pill-online";
  pillText.textContent = "Online";
  pill.querySelector(".dot").classList.remove("pulse");
  alertBox.style.display = "none";
  btnOpen.disabled = false;
  _online = true;
}

function setOffline() {
  pill.className = "pill pill-offline";
  pillText.textContent = "Offline";
  pill.querySelector(".dot").classList.remove("pulse");
  alertBox.style.display = "block";
  btnOpen.disabled = false;
  _online = false;
}

async function checkServer() {
  try {
    const ctrl  = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), 3000);
    await fetch(APP_URL, { method: "GET", mode: "no-cors", signal: ctrl.signal });
    clearTimeout(timer);
    setOnline();
  } catch {
    setOffline();
  }
}

btnOpen.addEventListener("click", () => {
  if (_online) {
    chrome.runtime.sendMessage({ action: "openApp" });
  } else {
    chrome.runtime.sendMessage({ action: "startAndOpen" });
  }
  window.close();
});

checkServer();
