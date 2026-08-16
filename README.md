# dcc-mcp-PowerPoint

PowerPoint adapter for the DCC-MCP ecosystem — the **application layer** over
[dcc-mcp-office](https://github.com/dcc-mcp/dcc-mcp-office): deck generation
from structured content, slide composition, review decks from DCC renders,
batch PDF conversion and text replacement (via the shared core), previews and
visual validation.

**Status:** M1 capability live. The designed pipeline runs end-to-end:
Deck IR (`office-ir/1.0`) → Open XML compile (python-pptx) → desktop COM
render (PowerPoint `DispatchEx`, dedicated instance) → PDF + per-slide PNG
previews → structural validation. Skill scripts are executable through the
dcc-mcp gateway (`DCC_MCP_POWERPOINT_SKILL_PATHS`). The shared C# host in
`dcc-mcp-office` remains the production COM path (M1 there).

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

## Agent usage (via dcc-mcp-cli / gateway)

1. Register the skill packs with the gateway:
   `set DCC_MCP_POWERPOINT_SKILL_PATHS=<repo>\src\dcc_mcp_powerpoint\skills`
2. Run a skill script directly (gateway `execute_script` contract —
   stdin JSON or CLI flags):
   ```bash
   python src/dcc_mcp_powerpoint/skills/powerpoint-deck/scripts/generate_deck.py \
     --input examples/dcc_mcp_framework_intro.json --out out
   ```
3. Validate the packs with the official linter:
   ```bash
   dcc-mcp-cli lint src/dcc_mcp_powerpoint/skills --warnings-as-errors --non-interactive
   ```

## Development

```bash
pip install -e .[dev,sidecar]
pytest
```

## License

MIT
