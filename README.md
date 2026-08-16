# dcc-mcp-PowerPoint

PowerPoint adapter for the DCC-MCP ecosystem — the **application layer** over
[dcc-mcp-office](https://github.com/dcc-mcp/dcc-mcp-office): deck generation
from structured content, slide composition, review decks from DCC renders,
batch PDF conversion and text replacement (via the shared core), previews and
visual validation.

**Status:** M0 scaffold. Protocol/IR and the C# `office-host` skeleton are
landing in `dcc-mcp-office`; this repo wires the PowerPoint application
semantics on top once M1 (COM MVP) ships there.

## Design (two-layer sidecar)

```text
Agent → dcc-mcp-gateway (Rust, dcc-mcp-core)
          → dcc-mcp-server sidecar  (lifecycle: RFC #998)
              → office-host.exe --app=powerpoint  (C#, dcc-mcp-office)
                  → POWERPNT.EXE (COM, interactive user session)
```

- Gateway + sidecar lifecycle: existing `dcc-mcp-core` machinery.
- `office-host`: shared C# runtime from `dcc-mcp-office`
  (download/verify/launch via `dcc-mcp-release-artifacts`).
- This repo owns: PowerPoint IR semantics, `powerpoint.deck.generate` /
  `powerpoint.slide.compose` / `powerpoint.slide.render`, template
  registry entries, and the PowerPoint skill packs.

## Skill packs

- `powerpoint-deck` — generate an editable deck from structured content
  (template → semantic layouts → IR → Open XML → COM finalize → previews →
  validation loop; proposal §15.3/§15.4).
- `powerpoint-review` — `dcc.review-deck-from-renders`: build review decks
  from Maya/Houdini/Blender/Unreal renders + shot/asset metadata
  (proposal §15.7).

## Development

```bash
pip install -e .[dev,sidecar]
pytest
```

## License

MIT
