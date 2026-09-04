# Install the PowerPoint adapter

The adapter is an out-of-process PowerPoint service. It does not install an
Office add-in and it never attaches to a user-owned PowerPoint process. The
official `dcc-mcp-server` owns MCP transport, FileRegistry registration and
heartbeats; this package owns PowerPoint skill packs and launches the shared
`dcc-office-host` for structured Office work.

## Requirements

- Windows 10/11 with an interactive user session.
- Desktop Microsoft PowerPoint for COM render/export operations. Open XML
  compile and structural inspection do not require PowerPoint.
- Current `dcc-mcp-cli` and `dcc-mcp-server` from the same DCC-MCP release.

Confirm the runtime before installing the adapter:

```powershell
dcc-mcp-cli doctor
dcc-mcp-server --version
```

## Released standalone bundle

Download `dcc-mcp-powerpoint-windows-x64.zip` and its `.sha256` file from the
same GitHub release. Verify the archive before extracting it:

```powershell
$expected = ((Get-Content -LiteralPath .\dcc-mcp-powerpoint-windows-x64.zip.sha256) -split '\s+')[0]
$actual = (Get-FileHash -Algorithm SHA256 -LiteralPath .\dcc-mcp-powerpoint-windows-x64.zip).Hash.ToLowerInvariant()
if ($actual -ne $expected) { throw "PowerPoint adapter checksum mismatch" }
Expand-Archive -LiteralPath .\dcc-mcp-powerpoint-windows-x64.zip -DestinationPath .\dcc-mcp-powerpoint
```

Start the adapter from the workspace it is allowed to read and write. The
launcher passes the bundled skill directory and its own Python runtime to the
official server, which then publishes a `powerpoint` registry row:

```powershell
Set-Location <workspace-root>
<adapter-root>\dcc-mcp-powerpoint.exe serve
dcc-mcp-cli wait-ready --dcc-type powerpoint
dcc-mcp-cli list
```

Keep the adapter process running while tools are in use. Stop it with Ctrl+C
or the guarded `dcc-mcp-cli stop-instance` workflow.

## Source checkout

For adapter development, use an isolated virtual environment:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e ".[dev,sidecar]"
.\.venv\Scripts\dcc-mcp-powerpoint.exe serve
```

## Verification

Registration is not proven by package installation. Verify the live route in
order, copying the tool slug returned by `search` rather than constructing it:

```powershell
dcc-mcp-cli doctor
dcc-mcp-cli list
dcc-mcp-cli wait-ready --dcc-type powerpoint
dcc-mcp-cli search --dcc-type powerpoint --query "inspect presentation"
dcc-mcp-cli describe <slug-from-search>
```

If `list` has no PowerPoint row, inspect the adapter process output and the log
paths reported by `dcc-mcp-cli doctor`. Do not fall back to generic desktop
automation.
