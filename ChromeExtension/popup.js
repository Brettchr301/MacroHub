// MacroHub Bridge popup

const BRIDGE_BASE = "http://localhost:9876";

function setStatus(msg, color) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.style.color = color || "#888";
}

function setPill(connected) {
  const pill = document.getElementById("connPill");
  pill.textContent = connected ? "Connected" : "Not connected";
  pill.className = "pill" + (connected ? "" : " off");
}

async function checkConnection() {
  try {
    const r = await fetch(BRIDGE_BASE + "/aiide/ping", { method: "GET" });
    if (!r.ok) {
      setPill(false);
      setStatus("MacroHub bridge responded with error status.", "#CF6F6F");
      return;
    }
    setPill(true);
    setStatus("MacroHub bridge is reachable on port 9876.", "#6FCF6F");
  } catch (_) {
    setPill(false);
    setStatus("MacroHub bridge not running. In AI IDE click Start Bridge.", "#CF6F6F");
  }
}

document.getElementById("checkBtn").addEventListener("click", checkConnection);

document.getElementById("captureBtn").addEventListener("click", () => {
  setStatus("Capturing latest response from active tab...", "#888");
  chrome.runtime.sendMessage({ action: "captureFromPopup" }, (resp) => {
    if (chrome.runtime.lastError) {
      setStatus("Extension error: " + chrome.runtime.lastError.message, "#CF6F6F");
      return;
    }
    if (resp && resp.ok === false) {
      setStatus(resp.error || "Capture failed.", "#CF6F6F");
      return;
    }
    setStatus("Captured and sent to MacroHub AI IDE.", "#6FCF6F");
  });
});

checkConnection();
