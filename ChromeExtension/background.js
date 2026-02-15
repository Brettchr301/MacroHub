// MacroHub Bridge service worker
// Receives popup messages and forwards to the active tab content script.

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.action !== "captureFromPopup") {
    return;
  }

  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    if (!tabs || !tabs.length || !tabs[0].id) {
      sendResponse({ ok: false, error: "No active tab available." });
      return;
    }
    chrome.tabs.sendMessage(tabs[0].id, { action: "capture" }, (response) => {
      if (chrome.runtime.lastError) {
        sendResponse({ ok: false, error: chrome.runtime.lastError.message });
        return;
      }
      sendResponse(response || { ok: true });
    });
  });

  return true;
});
