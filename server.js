// MCP server that bridges Claude Code -> (Tailscale HTTP) -> this process -> (localhost WS)
// -> Edge extension -> authenticated M365 Copilot web UI -> answer text -> back to Claude Code.
//
// Two listeners:
//   1. Streamable-HTTP MCP endpoint on the Tailscale interface (token-gated) for Claude Code.
//   2. A localhost-only WebSocket the paired Edge extension dials into.
//
// The extension is the only thing that touches the real, signed-in browser session, so all the
// auth/Conditional-Access/grounding concerns live there, not here.

import express from "express";
import { WebSocketServer } from "ws";
import { randomUUID } from "node:crypto";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { z } from "zod";

// ---- config ---------------------------------------------------------------
const cfg = {
  authToken: process.env.MCP_AUTH_TOKEN || "",
  bindAddr: process.env.MCP_BIND_ADDR || "127.0.0.1",
  port: Number(process.env.MCP_PORT || 7800),
  bridgeAddr: process.env.BRIDGE_BIND_ADDR || "127.0.0.1",
  bridgePort: Number(process.env.BRIDGE_PORT || 8765),
  defaultSurface: process.env.DEFAULT_SURFACE || "m365",
  askTimeoutMs: Number(process.env.ASK_TIMEOUT_MS || 120000),
};

if (!cfg.authToken || cfg.authToken === "change-me") {
  console.error("[fatal] set MCP_AUTH_TOKEN to a real secret (openssl rand -hex 32)");
  process.exit(1);
}

const log = (...a) => console.error(new Date().toISOString(), ...a);

// ---- extension bridge -----------------------------------------------------
// One extension connection at a time (one browser). Newest wins.
let extSocket = null;
const pending = new Map(); // id -> { resolve, reject, timer }

const wss = new WebSocketServer({ host: cfg.bridgeAddr, port: cfg.bridgePort });
wss.on("connection", (ws) => {
  log("[bridge] extension connected");
  if (extSocket && extSocket.readyState === 1) extSocket.close(4000, "superseded");
  extSocket = ws;
  ws.on("message", (buf) => {
    let msg;
    try { msg = JSON.parse(buf.toString()); } catch { return; }
    const p = pending.get(msg.id);
    if (!p) return;
    pending.delete(msg.id);
    clearTimeout(p.timer);
    if (msg.error) p.reject(new Error(msg.error));
    else p.resolve(msg.result);
  });
  ws.on("close", () => { if (extSocket === ws) extSocket = null; log("[bridge] extension disconnected"); });
  ws.on("error", (e) => log("[bridge] ws error", e.message));
});

function callExtension(type, payload, timeoutMs) {
  return new Promise((resolve, reject) => {
    if (!extSocket || extSocket.readyState !== 1) {
      return reject(new Error("Edge extension not connected. Open Edge on the laptop with the m365-copilot-mcp extension loaded."));
    }
    const id = randomUUID();
    const timer = setTimeout(() => {
      pending.delete(id);
      reject(new Error(`Timed out after ${timeoutMs}ms waiting for the Copilot UI.`));
    }, timeoutMs);
    pending.set(id, { resolve, reject, timer });
    extSocket.send(JSON.stringify({ id, type, ...payload }));
  });
}

// ---- MCP server -----------------------------------------------------------
function buildMcpServer() {
  const server = new McpServer({ name: "m365-copilot", version: "0.1.0" });

  server.registerTool(
    "ask_copilot",
    {
      title: "Ask M365 Copilot",
      description:
        "Send a prompt to the authenticated Microsoft 365 Copilot web UI on the work laptop and return its answer as plain text. Use for questions that benefit from the user's enterprise data (email, Teams, SharePoint, tenant docs) or fresh web grounding that this agent cannot access directly.",
      inputSchema: {
        prompt: z.string().min(1).describe("The natural-language question to put to Copilot."),
        surface: z
          .enum(["work", "m365", "web"])
          .optional()
          .describe("Which Copilot surface to use. work=enterprise grounding (m365.cloud.microsoft), m365=Copilot Chat (copilot.cloud.microsoft), web=consumer web grounding. Defaults to server config."),
        web_grounding: z
          .boolean()
          .optional()
          .describe("If false, ask Copilot to answer from work data only (single-turn toggle). Default true."),
        new_chat: z
          .boolean()
          .optional()
          .describe("Start a fresh conversation instead of continuing the current one. Default false."),
      },
    },
    async ({ prompt, surface, web_grounding, new_chat }) => {
      const result = await callExtension(
        "ask",
        {
          prompt,
          surface: surface || cfg.defaultSurface,
          webGrounding: web_grounding !== false,
          newChat: !!new_chat,
        },
        cfg.askTimeoutMs
      );
      const text = (result && result.text) ? result.text : "(no answer captured)";
      const cites = result && Array.isArray(result.citations) && result.citations.length
        ? "\n\nSources:\n" + result.citations.map((c, i) => `${i + 1}. ${c}`).join("\n")
        : "";
      return { content: [{ type: "text", text: text + cites }] };
    }
  );

  server.registerTool(
    "copilot_surfaces",
    {
      title: "List Copilot surfaces",
      description:
        "Report which M365 Copilot surfaces are currently open and authenticated in the laptop's Edge, and whether enterprise 'Work' grounding is available. Use this once to confirm entitlement before relying on a surface.",
      inputSchema: {},
    },
    async () => {
      const result = await callExtension("surfaces", {}, 20000);
      return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
    }
  );

  return server;
}

// ---- HTTP transport (Streamable HTTP, stateless) --------------------------
const app = express();
app.use(express.json({ limit: "2mb" }));

function checkAuth(req, res) {
  const hdr = req.headers["authorization"] || "";
  const token = hdr.startsWith("Bearer ") ? hdr.slice(7) : "";
  if (token !== cfg.authToken) {
    res.status(401).json({ jsonrpc: "2.0", error: { code: -32001, message: "Unauthorized" }, id: null });
    return false;
  }
  return true;
}

app.get("/healthz", (_req, res) => {
  res.json({ ok: true, extensionConnected: !!(extSocket && extSocket.readyState === 1) });
});

app.post("/mcp", async (req, res) => {
  if (!checkAuth(req, res)) return;
  const server = buildMcpServer();
  const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
  res.on("close", () => { transport.close(); server.close(); });
  try {
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (e) {
    log("[mcp] handler error", e);
    if (!res.headersSent) {
      res.status(500).json({ jsonrpc: "2.0", error: { code: -32603, message: "Internal error" }, id: null });
    }
  }
});

// Stateless server: GET/DELETE (SSE stream, session teardown) are not used.
app.get("/mcp", (_req, res) => res.status(405).json({ jsonrpc: "2.0", error: { code: -32000, message: "Method not allowed" }, id: null }));
app.delete("/mcp", (_req, res) => res.status(405).json({ jsonrpc: "2.0", error: { code: -32000, message: "Method not allowed" }, id: null }));

app.listen(cfg.port, cfg.bindAddr, () => {
  log(`[mcp] listening on http://${cfg.bindAddr}:${cfg.port}/mcp`);
  log(`[bridge] extension WebSocket on ws://${cfg.bridgeAddr}:${cfg.bridgePort}`);
  log(`[cfg] default surface=${cfg.defaultSurface} askTimeout=${cfg.askTimeoutMs}ms`);
});
