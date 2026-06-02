# m365-copilot-mcp

An MCP server that lets an MCP client (e.g. Claude Code) ask the **signed-in Microsoft 365 Copilot
web UI** running in your browser, and get the answer back as text. No Copilot API, no app
registration, no extra license — it drives the web surface you're already authenticated to, through
a small browser extension.

Useful when you have a Copilot seat in the browser but **no programmatic access** to the Copilot /
Microsoft Graph APIs (no add-on license, or no admin consent for the API scopes).

## How it works

```
MCP client  ──Streamable-HTTP MCP (token-gated)──►  server.js  ──ws://127.0.0.1──►  browser extension
                                                                                       │ content script
                                                                                       ▼ types prompt / scrapes answer
                                                                               M365 Copilot web (your session)
```

- **`server.js`** exposes the MCP endpoint over HTTP and a loopback WebSocket bridge.
- **`extension/`** is an MV3 (Chromium/Edge) extension. Its service worker dials the loopback bridge;
  its content script types your prompt into the Copilot tab and scrapes the streamed answer.

Only the extension touches the browser session, so there's no process injection, no remote-debugging
port and no second profile — it reuses your live SSO and stays low-profile to endpoint security.

## Why this shape

The only authenticated, policy-compliant Copilot session is the one in your interactive,
signed-in browser. So an extension *inside* that browser does the talking, and the Node bridge
exposes it to the MCP client. The Node process never touches the session — only the extension does —
so it can run in any OS session; only the extension needs the interactive desktop.

## Install

### 1. Server

```bash
npm install
cp .env.example .env      # then edit
```

Set in `.env` (or as real env vars):
- `MCP_AUTH_TOKEN` — a bearer the client must present. Generate: `openssl rand -hex 32`
- `MCP_BIND_ADDR` — bind address. Use your tailnet/VPN IP to reach it from another machine, or
  `127.0.0.1` for on-box only.
- `DEFAULT_SURFACE` — `work` / `m365` / `web` (see below).

Run it: `node server.js` (or, on Windows, `./run-server.ps1` after exporting `MCP_AUTH_TOKEN` or
placing a `.token` file beside the script — see *Always-on* below).

### 2. Extension

`edge://extensions` (or `chrome://extensions`) → enable **Developer mode** → **Load unpacked** →
select the `extension/` folder. Open a Copilot tab and stay signed in.

### 3. Register with your MCP client

For Claude Code:

```bash
claude mcp add --transport http m365-copilot http://<BIND_ADDR>:7800/mcp \
  --header "Authorization: Bearer <token>"
```

## Tools

- **`copilot_surfaces`** — report which Copilot surfaces are open + authenticated, and whether the
  enterprise **Work** grounding toggle is present. Run this first to confirm entitlement.
- **`ask_copilot { prompt, surface?, web_grounding?, new_chat? }`** — ask a question, get text back.

Surfaces:

| key  | URL                        | grounding |
|------|----------------------------|-----------|
| work | `m365.cloud.microsoft/chat`| enterprise (needs Copilot entitlement) |
| m365 | `copilot.cloud.microsoft`  | M365 Copilot Chat |
| web  | `copilot.microsoft.com`    | consumer / web only |

## Always-on (Windows)

`run-server.ps1` reads the token from `MCP_AUTH_TOKEN` or a `.token` file beside it (keep `.token`
out of git — it's gitignored), opens a firewall rule on the tailnet adapter, and starts the server.
Wire it as a logon Scheduled Task so it survives reboot:

```powershell
$pwsh = (Get-Command pwsh).Source
$action = New-ScheduledTaskAction -Execute $pwsh -Argument '-NoProfile -ExecutionPolicy Bypass -File "C:\path\to\run-server.ps1"'
$trigger = New-ScheduledTaskTrigger -AtLogOn
Register-ScheduledTask -TaskName 'm365-copilot-mcp' -Action $action -Trigger $trigger -RunLevel Highest -Force
```

## Tuning the selectors

These web apps ship obfuscated, shifting DOM. `extension/content.js` uses per-host selector *hints*
with generic fallbacks and a stability heuristic (an answer is "done" after ~2.5 s with no text
growth). If a surface returns `composer not found` or empty answers:

1. Run `copilot_surfaces` — the probe output reports what was found (composer tag, send button,
   message count, auth state).
2. Inspect the composer / latest answer bubble in your browser and update the `HINTS[<host>]`
   arrays in `content.js`.
3. Reload the extension and retry.

## Security & caveats

- The MCP endpoint requires a bearer token. Bind it to a private interface (tailnet/VPN) and/or
  restrict it with a host firewall and ACL. The extension bridge is **loopback only** — never
  expose it.
- This **automates the Copilot UI under your own identity**. That can be a grey area under
  Microsoft's service terms and may draw attention from DLP / security tooling, especially on a
  managed/enterprise tenant. Understand your org's policies before using it there.
- Answers are AI-generated and may be wrong — verify before acting.

## Known limitations

- One browser / one extension connection at a time (newest wins).
- Long answers near the timeout can fail; raise `ASK_TIMEOUT_MS`.
- MV3 service workers get evicted when idle; an alarm-based reconnect (`alarms` permission) keeps
  the bridge healthy, but a dropped bridge may take up to ~30 s to recover.
- Selectors break when the vendor reships the UI; the probe/hints workflow above is the fix.

## License

MIT
