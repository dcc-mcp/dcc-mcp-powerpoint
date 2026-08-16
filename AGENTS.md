# AGENTS.md — dcc-mcp-PowerPoint Agent Navigation Map

> Progressive disclosure: this file is a **map**, not an encyclopedia.

## 30-Second Summary

`dcc-mcp-PowerPoint` is the **thin PowerPoint adapter** over
`dcc-mcp-office`. It owns PowerPoint application semantics only: deck
generation, slide composition, review decks, PowerPoint skill packs and the
office-host launcher. Shared machinery (protocol, IR envelope, C# COM
runtime, Open XML worker, jobs, security policy) comes from
`dcc-mcp-office` + `dcc-mcp-core`.

**Current status:** M1 capability live — deck pipeline runs end-to-end
(Deck IR → Open XML compile → COM render → PDF/previews → validation),
skill scripts executable via the gateway. The shared C# host remains the
production COM path (dcc-mcp-office M1).

## Repo Map

| Path | What it is |
|---|---|
| `src/dcc_mcp_powerpoint/` | adapter package |
| `src/dcc_mcp_powerpoint/deck_ir.py` | Deck IR contract (mirrors dcc-mcp-office-ir) |
| `src/dcc_mcp_powerpoint/compiler.py` | Open XML compiler: semantic-layout registry → PPTX |
| `src/dcc_mcp_powerpoint/render.py` | COM renderer: `DispatchEx` dedicated instance → PDF + PNGs |
| `src/dcc_mcp_powerpoint/validate.py` | structural + artifact validation reports |
| `src/dcc_mcp_powerpoint/sidecar/office_host.py` | launcher: locate/verify/start `office-host.exe --app=powerpoint` |
| `src/dcc_mcp_powerpoint/skills/powerpoint-deck/` | SKILL.md + tools.yaml + scripts (generate/validate/render) |
| `src/dcc_mcp_powerpoint/skills/powerpoint-review/` | `review_deck_from_renders` skill |
| `examples/` | framework-intro Deck IR + generated PPTX/PDF/previews |
| `tests/` | pytest (COM only via subprocess boundary) |
| `docs/adr/` | adapter-level decisions |

## Upstream Dependencies

- `dcc-mcp-core` (pip) — gateway, skills runtime, sidecar lifecycle.
- `dcc-mcp-office` — Rust crates (`dcc-mcp-office-protocol`,
  `dcc-mcp-office-ir`, `dcc-mcp-office-tools`, `dcc-mcp-office-security`)
  + the `office-host` C# binaries distributed via
  `dcc-mcp-release-artifacts`.
- `dcc-mcp-computer-use` — UIA/visual fallback is reused, not rebuilt.

## Capabilities owned here (proposal §11.2)

- `powerpoint.deck.generate` — requirement → outline → template/layout →
  Deck IR → Open XML compile → COM finalize → previews → validation loop.
- `powerpoint.slide.compose` — compose one slide from semantic layout +
  content blocks.
- `powerpoint.slide.render` — render slides to PNG previews.
- `powerpoint.animation.apply` / `powerpoint.slideshow.control` — later.

## Test

```bash
pip install -e .[dev]
pytest
ruff check src tests
```

Engineering agreement: [CONTRIBUTING.md](./CONTRIBUTING.md) (plus the shared
rules in `dcc-mcp-office/CONTRIBUTING.md`).

Golden files / visual snapshots for PowerPoint live in
`dcc-mcp-office/tests/` (shared matrix); this repo adds PowerPoint-specific
cases as they land.
