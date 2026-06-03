// Runs inside the authenticated Copilot tab. Types the prompt, waits for the streamed answer
// to settle, and returns plain text. Selectors are best-effort with generic fallbacks; the
// `probe` command dumps what it found so they can be tuned against the live DOM.

(() => {
  if (window.__m365CopilotBridgeLoaded) return;
  window.__m365CopilotBridgeLoaded = true;

  const STABLE_MS = 4000;     // answer considered done after this long with no text growth
  const POLL_MS = 500;
  const MAX_WAIT_MS = 115000; // sit just under the server's 120s ASK_TIMEOUT_MS

  // Transient "thinking/grounding" status text Copilot parks on while it searches M365/web.
  // These can stay static for >STABLE_MS, so they must never be accepted as the final answer.
  const PLACEHOLDER_RE = /^(generating( a)? response|lining things up|crafting search queries|searching|looking (through|across|into)|working on it|thinking|reasoning|getting (things|everything) ready|i'?m (planning|looking|searching|gathering)|let me|one moment|just a (sec|second|moment)|hang on|preparing|analy[sz]ing|reviewing|checking)/i;

  // Citation/sources/footer blocks that are NOT the answer body.
  const NONANSWER_RE = /^(sources?|references?|related|see also|follow[- ]?up|suggested|you might|ai-generated|copilot)\b/i;

  // Per-host selector hints. Empty arrays fall back to heuristics. Tune these after `probe`.
  // `messages` is ordered most-specific -> most-generic; the first selector that matches any
  // node wins. The generic `[role="listitem"]` is a last resort and is filtered hard in
  // pickAnswer() so trailing Sources / suggestion items are never mistaken for the answer.
  const HINTS = {
    "m365.cloud.microsoft": {
      composer: ['div[contenteditable="true"]', 'span[contenteditable="true"]', 'textarea'],
      send: ['button[aria-label*="Send" i]', 'button[title*="Send" i]'],
      // Verified live 2026-06-03: assistant answer body is div.fai-CopilotMessage__content
      // (excludes the "Copilot said:" heading, the name, and the trailing __footnote "Sources").
      // Generic fallbacks kept for UI drift.
      messages: ['div[class*="CopilotMessage__content" i]', 'div[class*="CopilotMessage" i]:not([class*="UserMessage" i])', '[data-content="message-body"]', '[class*="responseMessage" i]', '[class*="botMessage" i]'],
      groundingToggle: ['button[aria-label*="Work" i]', 'button[aria-label*="Web" i]'],
    },
    "copilot.cloud.microsoft": {
      composer: ['div[contenteditable="true"]', 'span[contenteditable="true"]', 'textarea'],
      send: ['button[aria-label*="Send" i]', 'button[title*="Send" i]'],
      messages: ['div[class*="CopilotMessage__content" i]', 'div[class*="CopilotMessage" i]:not([class*="UserMessage" i])', '[data-content="message-body"]', '[class*="responseMessage" i]', '[class*="botMessage" i]'],
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

  // True while Copilot is still streaming: it swaps the Send control for a Stop/cancel control.
  // Best-effort — pickAnswer()'s placeholder filtering is the safety net if this misses.
  function isGenerating() {
    return [...document.querySelectorAll('button, [role="button"]')].some((b) => {
      const t = (b.getAttribute("aria-label") || b.title || b.textContent || "").toLowerCase();
      return /\b(stop (response|responding|generating|streaming)|cancel)\b/.test(t) && isVisible(b);
    });
  }

  // Pick the assistant's answer node for the *current* turn. Excludes the echoed user prompt,
  // pure Sources/suggestion blocks, transient thinking placeholders, and any node that already
  // existed before this turn (the `exclude` set) — so a longer *prior* answer in the same thread
  // can't win. Returns the last remaining candidate in document order (the newest turn). This
  // replaces both the old nodes[last] grab (locked onto "Sources") and the largest-node heuristic
  // (locked onto a longer previous answer when new_chat didn't reset the thread).
  function pickAnswer(prompt, exclude) {
    const promptHead = (prompt || "").trim().slice(0, 50).toLowerCase();
    const cands = messageNodes().filter((n) => {
      if (exclude && exclude.has(n)) return false;                              // pre-existing turn
      const txt = (n.innerText || "").trim();
      if (!txt) return false;
      if (promptHead && txt.toLowerCase().startsWith(promptHead)) return false; // user echo
      if (/you said:/i.test(txt.slice(0, 60))) return false;                    // user turn / whole-turn wrapper
      if (NONANSWER_RE.test(txt) && txt.length < 240) return false;             // sources/suggestions
      if (PLACEHOLDER_RE.test(txt) && txt.length < 160) return false;           // thinking placeholder
      return true;
    });
    return cands.length ? cands[cands.length - 1] : null;                       // newest turn
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

    // Snapshot the answer nodes that already exist so a prior turn in the same thread is never
    // mistaken for this turn's answer (matters when new_chat didn't fully reset the conversation).
    const before = new Set(messageNodes());
    setComposerText(composer, prompt);
    await sleep(150);
    await submit(composer);

    // Wait for the turn to start: either a brand-new answer node appears or the Stop control shows.
    const start = Date.now();
    while (Date.now() - start < 20000) {
      if (messageNodes().some((n) => !before.has(n)) || isGenerating()) break;
      await sleep(POLL_MS);
    }

    // Settle the answer. Three conditions must all hold for STABLE_MS before we accept:
    //   1. not generating (no Stop control),
    //   2. the chosen answer node holds real (non-placeholder) text,
    //   3. that text has stopped growing.
    // This kills both prior failure modes: parking on a "Generating…" placeholder (fails #1/#2)
    // and capturing the trailing "Sources" block (pickAnswer never selects it).
    let target = null;
    let lastText = "";
    let stableSince = null;
    while (Date.now() - start < MAX_WAIT_MS) {
      const ans = pickAnswer(prompt, before);
      const txt = ans ? (ans.innerText || "").trim() : "";
      const real = txt.length > 0 && !(PLACEHOLDER_RE.test(txt) && txt.length < 160);
      if (ans) target = ans;

      if (isGenerating() || !real || txt !== lastText) {
        lastText = txt;
        stableSince = null;
      } else if (stableSince === null) {
        stableSince = Date.now();
      } else if (Date.now() - stableSince > STABLE_MS) {
        break;
      }
      await sleep(POLL_MS);
    }

    if (!target || !lastText) {
      throw new Error("no answer text captured (selectors may be stale; run probe to inspect the DOM)");
    }

    // Citations can live in the answer node or a sibling Sources block; scan both.
    const cites = new Set(collectCitations(target));
    messageNodes().forEach((n) => {
      if (NONANSWER_RE.test((n.innerText || "").trim())) collectCitations(n).forEach((u) => cites.add(u));
    });
    return { text: lastText, citations: [...cites].slice(0, 20) };
  }

  function probe() {
    const composer = findComposer();
    const send = findSendButton();
    const msgs = messageNodes();
    const signedOut = /sign in|sign-in|log in/i.test(document.body.innerText.slice(0, 4000)) && msgs.length === 0;
    const grounding = hints().groundingToggle.some((s) => document.querySelector(s));
    const ans = pickAnswer("");
    // Compact DOM sample for tuning selectors against the live page without redeploying.
    const sample = msgs.slice(-8).map((n) => ({
      tag: n.tagName,
      role: n.getAttribute("role") || null,
      cls: (typeof n.className === "string" ? n.className : "").slice(0, 60),
      data: Object.keys(n.dataset || {}).join(",") || null,
      len: (n.innerText || "").trim().length,
      snip: (n.innerText || "").trim().slice(0, 80),
    }));
    const buttons = [...document.querySelectorAll('button, [role="button"]')]
      .map((b) => (b.getAttribute("aria-label") || b.title || b.textContent || "").trim())
      .filter((t) => t && t.length < 40)
      .slice(0, 30);
    return {
      url: location.href,
      authenticated: !signedOut,
      composerFound: !!composer,
      composerTag: composer ? (composer.tagName + (composer.getAttribute("contenteditable") ? "[ce]" : "")) : null,
      sendButtonFound: !!send,
      messageCount: msgs.length,
      groundingToggleFound: grounding,
      generating: isGenerating(),
      answerLen: ans ? (ans.innerText || "").trim().length : 0,
      answerSnip: ans ? (ans.innerText || "").trim().slice(0, 120) : null,
      nodes: sample,
      buttons,
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
