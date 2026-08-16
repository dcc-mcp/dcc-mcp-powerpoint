# AGENTS.md — dcc-mcp-PowerPoint Agent Navigation Map

> Progressive disclosure: this file is a **map**, not an encyclopedia.

## 30-Second Summary

`dcc-mcp-PowerPoint` is the **thin PowerPoint adapter** over
`dcc-mcp-office`. It owns PowerPoint application semantics only: deck
generation, slide composition, review decks, PowerPoint skill packs and the
office-host launcher. Shared machinery (protocol, IR envelope, C# COM
runtime, Open XML worker, jobs, security policy) comes from
`dcc-mcp-office` + `dcc-mcp-core`.

**Current status:** M0 scaffold — package + skill packs + launcher stub.
No COM wiring yet (that lands in dcc-mcp-office M1).

## Repo Map

| Path | What it is |
|---|---|
| `src/dcc_mcp_powerpoint/` | adapter package |
| `src/dcc_mcp_powerpoint/sidecar/office_host.py` | launcher: locate/verify/start `office-host.exe --app=powerpoint` |
| `src/dcc_mcp_powerpoint/skills/powerpoint-deck/SKILL.md` | deck generation workflow (proposal §15.3/§15.4) |
| `src/dcc_mcp_powerpoint/skills/powerpoint-review/SKILL.md` | `dcc.review-deck-from-renders` (proposal §15.7) |
| `tests/` | pytest (no Office needed for M0) |
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
