---
name: powerpoint-deck
description: >-
  Generate an editable PowerPoint deck from structured content following the
  DCC-MCP designed pipeline: Deck IR (office-ir/1.0) → Open XML compile →
  desktop COM render (PDF + per-slide PNG previews) → structural validation
  report. Use whenever the agent must produce a PPTX deck from a deck spec.
license: MIT
allowed-tools: Bash Read
metadata:
  dcc-mcp:
    dcc: powerpoint
    layer: domain
    stage: authoring
    version: 0.1.0
    tags:
      - powerpoint
      - deck
      - generate
      - pptx
      - pdf
    search-hint: >-
      generate deck, make ppt, create slides, powerpoint presentation,
      export pdf, slide previews, deck validation
    tools: tools.yaml
---

# powerpoint-deck (Authoring stage)

Deck generation through the designed pipeline (proposal §15.3/§15.4):

1. content planner picks semantic layouts (never raw coordinates)
2. Deck IR document (contract: dcc-mcp-office-ir presentation schema)
3. Open XML compiler builds the base PPTX (fast, Office-free)
4. desktop COM renderer exports PDF + per-slide previews (skipped with an
   explicit reason when PowerPoint is unavailable — never silent)
5. structural validation report

## Input contract

- `input` — path to a Deck IR JSON envelope
  (`schema_version: office-ir/1.0`, `kind: presentation`)
- `output_dir` — artifact directory
- `render` / `previews` — whether to run the COM render step

## Scripts

- `generate_deck` — IR → PPTX (+ PDF/previews + validation report)
- `validate_deck` — validate a Deck IR / artifacts without generating
- `render_deck` — render an existing PPTX to PDF + previews via COM

## Validation rules

- envelope contract enforced at load (`deck_ir`): unknown layouts, missing
  keys and bad types are hard errors with a json-path hint
- structural checks: schema version, slide count, titles, bullet budgets
- artifacts must exist and be non-empty

## Agent-visible summary

Result context: artifacts (pptx/pdf/previews), backend per step, validation
checks + warnings, and whether human review is recommended.
