# Launch the MCP server on the host machine. Loopback bridge + tailnet-only MCP endpoint.
# Run once after login, or wire as a Scheduled Task (logon trigger). The Node process does NOT
# need the interactive desktop — only the Edge extension does — so this may run in any session.

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

# --- config (override via real env vars / a secrets store; do not commit a token) ----------
# Token precedence: env var, else the ACL-restricted .token file holding your bearer token.
# The Scheduled Task relies on the file path.
if (-not $env:MCP_AUTH_TOKEN) {
  $tf = Join-Path $here ".token"
  if (Test-Path $tf) { $env:MCP_AUTH_TOKEN = (Get-Content $tf -Raw).Trim() }
  else { throw "No MCP_AUTH_TOKEN env and no .token file beside the script." }
}
$env:MCP_BIND_ADDR   = $env:MCP_BIND_ADDR   ?? "127.0.0.1"   # set to your tailnet IP to expose over Tailscale
$env:MCP_PORT        = $env:MCP_PORT        ?? "7800"
$env:BRIDGE_BIND_ADDR = "127.0.0.1"
$env:BRIDGE_PORT     = $env:BRIDGE_PORT     ?? "8765"
$env:DEFAULT_SURFACE = $env:DEFAULT_SURFACE ?? "work"   # m365.cloud.microsoft: entitled + Work grounding

# Allow inbound on the MCP port for the Tailscale adapter only (idempotent).
$ruleName = "m365-copilot-mcp $($env:MCP_PORT)"
if (-not (Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue)) {
  New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Action Allow `
    -Protocol TCP -LocalPort $env:MCP_PORT -InterfaceAlias "Tailscale" | Out-Null
  Write-Host "Added firewall rule '$ruleName' on Tailscale adapter"
}

Write-Host "Starting MCP server on http://$($env:MCP_BIND_ADDR):$($env:MCP_PORT)/mcp"
node server.js
