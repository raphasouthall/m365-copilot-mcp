// Runs inside the authenticated Copilot tab. Types the prompt, waits for the streamed answer
// to settle, and returns plain text. Selectors are best-effort with generic fallbacks; the
// `probe` command dumps what it found so they can be tuned against the live DOM.

(() => {
  if (window.__m365CopilotBridgeLoaded) return;
  window.__m365CopilotBridgeLoaded = true;

  const STABLE_MS = 2500;     // answer considered done after this long with no text growth
  const POLL_MS = 400;
  const MAX_WAIT_MS = 110000;

  // Per-host selector hints. Empty arrays fall back to heuristics. Tune these after `probe`.
  const HINTS = {
    "m365.cloud.microsoft": {
      composer: ['div[contenteditable="true"]', 'textarea'],
      send: ['button[aria-label*="Send" i]', 'button[title*="Send" i]'],
      messages: ['[data-content="message-body"]', '[class*="responseMessage"]', '[role="listitem"]'],
      groundingToggle: ['button[aria-label*="Work" i]', 'button[aria-label*="Web" i]'],
    },
    "copilot.cloud.microsoft": {
      composer: ['div[contenteditable="true"]', 'textarea'],
      send: ['button[aria-label*="Send" i]', 'button[title*="Send" i]'],
      messages: ['[data-content="message-body"]', '[class*="responseMessage"]', '[role="listitem"]'],
      groundingToggle: ['button[aria-label*="Work" i]', 'button[aria-label*="Web" i]'],
    },
    "copilot.microsoft.com": {
      composer: ['textarea', 'div[contenteditable="true"]'],
      send: ['button[aria-label*="Submit" i]', 'button[aria-label*="Send" i]'],
      messages: ['[data-content="ai-message"]', '[class*="message"]'],
      groundingToggle: [],
    },
  };

  const hints = () => HINTS[location.host] || { composer: [], send: [], messages: [], groundingToggle: [] };
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  function firstMatch(selectors, fallback) {
    for (const sel of selectors) {
      const el = document.querySelector(sel);
      if (el) return el;
    }
    return fallback ? fallback() : null;
  }

  function findComposer() {
    return firstMatch(hints().composer, () => {
      // heuristic: largest visible editable
      const cands = [
        ...document.querySelectorAll('div[contenteditable="true"], textarea, [role="textbox"]'),
      ].filter(isVisible);
      cands.sort((a, b) => area(b) - area(a));
      return cands[0] || null;
    });
  }

  function findSendButton() {
    return firstMatch(hints().send, () => {
      return [...document.querySelectorAll("button")].find((b) => {
        const t = (b.getAttribute("aria-label") || b.title || b.textContent || "").toLowerCase();
        return /send|submit/.test(t) && isVisible(b);
      });
    });
  }

  function messageNodes() {
    for (const sel of hints().messages) {
      const nodes = [...document.querySelectorAll(sel)].filter(isVisible);
      if (nodes.length) return nodes;
    }
    // heuristic fallback: list items / article-ish blocks with meaningful text
    return [...document.querySelectorAll('[role="listitem"], article, [class*="message" i]')]
      .filter((n) => isVisible(n) && n.innerText && n.innerText.trim().length > 0);
  }

  const isVisible = (el) => {
    if (!el) return false;
    const r = el.getBoundingClientRect();
    const s = getComputedStyle(el);
    return r.width > 0 && r.height > 0 && s.visibility !== "hidden" && s.display !== "none";
  };
  const area = (el) => { const r = el.getBoundingClientRect(); return r.width * r.height; };

  function setComposerText(el, text) {
    el.focus();
    if (el.tagName === "TEXTAREA" || el.tagName === "INPUT") {
      const setter = Object.getOwnPropertyDescriptor(
        el.tagName === "TEXTAREA" ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype,
        "value"
      ).set;
      setter.call(el, text);
      el.dispatchEvent(new Event("input", { bubbles: true }));
    } else {
      // contenteditable: clear, then insert via execCommand so the framework sees real input events
      const sel = window.getSelection();
      sel.removeAllRanges();
      const range = document.createRange();
      range.selectNodeContents(el);
      sel.addRange(range);
      document.execCommand("insertText", false, "");
      document.execCommand("insertText", false, text);
      el.dispatchEvent(new InputEvent("input", { bubbles: true, data: text, inputType: "insertText" }));
    }
  }

  async function submit(el) {
    const btn = findSendButton();
    if (btn && !btn.disabled) { btn.click(); return; }
    // fallback: Enter
    el.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, key: "Enter", code: "Enter", keyCode: 13 }));
    el.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true, key: "Enter", code: "Enter", keyCode: 13 }));
  }

  // Click the "new chat" control and wait for the thread to actually reset, then for a usable
  // composer to reappear. Bounded so it can never hang the surrounding ask: if no button matches
  // we return immediately and the ask continues in the current conversation (degraded, not stuck).
  async function startNewChat() {
    const findBtn = () =>
      [...document.querySelectorAll('button, a[role="button"], [role="button"]')].find((b) => {
        const t = (b.getAttribute("aria-label") || b.title || b.textContent || "").toLowerCase();
        return /new chat|new conversation|new topic|start new/.test(t) && isVisible(b);
      });
    const btn = findBtn();
    if (!btn) return; // no control found — stay in the current chat rather than block
    const before = messageNodes().length;
    btn.click();
    // SPA navigations re-render asynchronously; wait (bounded) for the thread to clear and the
    // composer to rehydrate before the caller types into it.
    const start = Date.now();
    while (Date.now() - start < 8000) {
      await sleep(300);
      const composer = findComposer();
      const count = messageNodes().length;
      if (composer && isVisible(composer) && (count === 0 || count < before)) return;
    }
  }

  function collectCitations(node) {
    if (!node) return [];
    const urls = new Set();
    node.querySelectorAll('a[href^="http"]').forEach((a) => urls.add(a.href));
    return [...urls].slice(0, 20);
  }

  async function ask({ prompt, newChat }) {
    if (newChat) await startNewChat();
    const composer = findComposer();
    if (!composer) throw new Error("composer not found on " + location.host + " (run probe to tune selectors)");

    const before = messageNodes().length;
    setComposerText(composer, prompt);
    await sleep(150);
    await submit(composer);

    // wait for a new message node to appear
    const start = Date.now();
    let target = null;
    while (Date.now() - start < MAX_WAIT_MS) {
      const nodes = messageNodes();
      if (nodes.length > before) { target = nodes[nodes.length - 1]; break; }
      await sleep(POLL_MS);
    }
    if (!target) throw new Error("no response element appeared (selectors may be stale)");

    // wait for streamed text to stabilise
    let lastText = "";
    let lastChange = Date.now();
    while (Date.now() - start < MAX_WAIT_MS) {
      const nodes = messageNodes();
      target = nodes[nodes.length - 1] || target;
      const txt = (target.innerText || "").trim();
      if (txt !== lastText) { lastText = txt; lastChange = Date.now(); }
      else if (lastText && Date.now() - lastChange > STABLE_MS) break;
      await sleep(POLL_MS);
    }

    return { text: lastText, citations: collectCitations(target) };
  }

  function probe() {
    const composer = findComposer();
    const send = findSendButton();
    const msgs = messageNodes();
    const signedOut = /sign in|sign-in|log in/i.test(document.body.innerText.slice(0, 4000)) && msgs.length === 0;
    const grounding = hints().groundingToggle.some((s) => document.querySelector(s));
    return {
      url: location.href,
      authenticated: !signedOut,
      composerFound: !!composer,
      composerTag: composer ? (composer.tagName + (composer.getAttribute("contenteditable") ? "[ce]" : "")) : null,
      sendButtonFound: !!send,
      messageCount: msgs.length,
      groundingToggleFound: grounding,
    };
  }

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    (async () => {
      try {
        if (msg.cmd === "ask") sendResponse(await ask(msg));
        else if (msg.cmd === "probe") sendResponse(probe());
        else sendResponse({ error: "unknown cmd" });
      } catch (e) {
        sendResponse({ error: String(e && e.message ? e.message : e) });
      }
    })();
    return true; // async response
  });
})();
