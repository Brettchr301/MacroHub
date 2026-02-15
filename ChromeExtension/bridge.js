// MacroHub Bridge content script
// Passive mode: capture completed responses and send to /aiide.
// Active mode: poll /prompt, inject prompt, and post /result.

(function () {
  "use strict";

  const BASE_URL = "http://localhost:9876";
  const RESULT_URL = BASE_URL + "/result";
  const PROMPT_URL = BASE_URL + "/prompt";
  const FILE_URL = BASE_URL + "/file";
  const AIIDE_URL = BASE_URL + "/aiide";

  let lastSentText = "";
  let captureDebounce = null;
  let injecting = false;

  function getLatestResponse() {
    const h = location.hostname;

    if (h.includes("copilot.microsoft.com") || h.includes("bing.com")) {
      const selectors = [
        '[data-testid="chat-message-assistant"]',
        'cib-message-group[source="bot"] cib-message',
        '[data-automationid="chat-answer-container"]',
        '.cib-chat-main-content [class*="responseText"]'
      ];
      for (const sel of selectors) {
        try {
          const nodes = document.querySelectorAll(sel);
          if (nodes && nodes.length) {
            const txt = (nodes[nodes.length - 1].innerText || "").trim();
            if (txt) return txt;
          }
        } catch (_) {}
      }
      return null;
    }

    if (h.includes("claude.ai")) {
      const nodes = document.querySelectorAll('[data-is-streaming="false"] .font-claude-message');
      if (!nodes.length) return null;
      return (nodes[nodes.length - 1].innerText || "").trim();
    }

    if (h.includes("chat.openai.com") || h.includes("chatgpt.com")) {
      const nodes = document.querySelectorAll('[data-message-author-role="assistant"]');
      if (!nodes.length) return null;
      return (nodes[nodes.length - 1].innerText || "").trim();
    }

    return null;
  }

  function findInput() {
    const selectors = [
      'textarea[data-testid="copilot-chat-input"]',
      'textarea[data-testid="chat-input"]',
      "cib-text-input textarea",
      'div[contenteditable="true"][role="textbox"]',
      "textarea[placeholder]"
    ];
    for (const sel of selectors) {
      try {
        const el = document.querySelector(sel);
        if (el && el.offsetParent !== null) return el;
      } catch (_) {}
    }
    return null;
  }

  function findSubmitButton() {
    const selectors = [
      'button[data-testid="submit-button"]',
      'button[aria-label*="Send"]',
      'button[aria-label*="Submit"]',
      'cib-text-input button[type="submit"]',
      'button[type="submit"]'
    ];
    for (const sel of selectors) {
      try {
        const btn = document.querySelector(sel);
        if (btn && !btn.disabled && btn.offsetParent !== null) return btn;
      } catch (_) {}
    }
    return null;
  }

  function setInputValue(el, text) {
    try {
      const proto = el.tagName === "TEXTAREA"
        ? window.HTMLTextAreaElement.prototype
        : window.HTMLInputElement.prototype;
      const desc = Object.getOwnPropertyDescriptor(proto, "value");
      if (desc && desc.set) {
        desc.set.call(el, text);
      } else {
        el.value = text;
      }
    } catch (_) {
      el.value = text;
    }
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
  }

  function injectPrompt(promptText) {
    const input = findInput();
    if (!input) return false;

    setInputValue(input, promptText);
    input.focus();

    setTimeout(() => {
      const btn = findSubmitButton();
      if (btn) {
        btn.click();
      } else {
        input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", keyCode: 13, bubbles: true }));
        input.dispatchEvent(new KeyboardEvent("keyup", { key: "Enter", keyCode: 13, bubbles: true }));
      }
    }, 300);

    return true;
  }

  function postResult(text, url) {
    if (!text) return;
    fetch(url || AIIDE_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ source: location.hostname, response: text, ts: Date.now() })
    }).catch(() => {});
  }

  function passiveCaptureTick() {
    if (injecting) return;
    const latest = getLatestResponse();
    if (latest && latest !== lastSentText) {
      lastSentText = latest;
      postResult(latest, AIIDE_URL);
    }
  }

  const observer = new MutationObserver(() => {
    clearTimeout(captureDebounce);
    captureDebounce = setTimeout(passiveCaptureTick, 1200);
  });

  if (document.body) {
    observer.observe(document.body, { childList: true, subtree: true });
  }

  function doInject(promptText) {
    if (!promptText) return;

    injecting = true;
    const before = getLatestResponse();

    if (!injectPrompt(promptText)) {
      injecting = false;
      postResult("ERROR: Could not find AI input box on this page.", RESULT_URL);
      return;
    }

    const deadline = Date.now() + 90000;
    const timer = setInterval(() => {
      const latest = getLatestResponse();
      const isNew = latest && latest !== before && latest !== lastSentText;
      const isStreaming = !!document.querySelector(
        '[class*="typing"], [class*="streaming"], cib-typing-indicator, [aria-label*="typing" i]'
      );

      if (isNew && !isStreaming) {
        clearInterval(timer);
        injecting = false;
        lastSentText = latest;
        postResult(latest, RESULT_URL);
      } else if (Date.now() > deadline) {
        clearInterval(timer);
        injecting = false;
        postResult("ERROR: Timed out waiting for AI response.", RESULT_URL);
      }
    }, 800);
  }

  function pollForPrompt() {
    if (injecting) return;

    fetch(PROMPT_URL, { method: "GET" })
      .then((r) => r.text())
      .then((prompt) => {
        const req = (prompt || "").trim();
        if (!req) return;

        if (req === "__FILE__") {
          fetch(FILE_URL, { method: "GET" })
            .then((r) => r.text())
            .then((txt) => doInject((txt || "").trim()))
            .catch(() => { injecting = false; });
        } else {
          doInject(req);
        }
      })
      .catch(() => {});
  }

  setInterval(pollForPrompt, 2000);

  chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (msg.action !== "capture") return;
    const latest = getLatestResponse();
    if (latest) {
      postResult(latest, AIIDE_URL);
      sendResponse({ ok: true });
    } else {
      sendResponse({ ok: false, error: "No assistant response found on this page yet." });
    }
  });
})();
