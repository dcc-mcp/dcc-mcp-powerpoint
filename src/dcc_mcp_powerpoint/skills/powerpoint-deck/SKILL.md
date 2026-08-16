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
    version: 0.2.0
    tags:
      - powerpoint
      - deck
      - generate
      - pptx
      - pdf
      - layers
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

Every generated deck is semantic-layer tagged (see `powerpoint-layers`):
background / decoration / header / content / accent / footer. Layers let
later edits hide decoration, rebrand accents and restack z-order without
touching content.

## Related skills

- `powerpoint-edit` — modify an existing deck with a ppt-patch/1.0 document
- `powerpoint-layers` — show/hide, restack and recolor semantic layers
- `powerpoint-plugins` — invoke project-specific deck transforms

Full per-layout Deck IR examples: `references/RECIPES.md`.

## Input contract

- `input` — path to a Deck IR JSON envelope
  (`schema_version: office-ir/1.0`, `kind: presentation`)
- `output_dir` — artifact directory
- `render` / `previews` — whether to run the COM render step

## Decision rules (document-pptx learnings)

- start from the audience and the decision the deck must support; one
  takeaway per slide
- brand fidelity matters → use the brand:// template registry, never guess
  brand colors or fonts (the official lockup resolves from the registry)
- narrative before code: build the slide outline, then the Deck IR, then
  compile — never author slide coordinates first
- resolve layouts by name (semantic_layout), never by guessed index

## Scripts

- `generate_deck` — IR → PPTX (+ PDF/previews + validation report)
- `validate_deck` — validate a Deck IR / artifacts without generating
- `render_deck` — render an existing PPTX to PDF + previews via COM
- `inventory_deck` — read-only PPTX inventory (notes coverage, alt-text
  gaps, shape counts) before editing or repair

## Validation rules

- envelope contract enforced at load (`deck_ir`): unknown layouts, missing
  keys and bad types are hard errors with a json-path hint
- structural checks: schema version, slide count, titles, bullet budgets
- every picture carries alt text (accessibility is authoring, not
  afterthought); `inventory_deck` reports gaps
- artifacts must exist and be non-empty

## Known limits

- python-pptx has no animation/transition API and no gradient-fill API;
  both need direct OOXML edits — never promise motion without it
- no true combo charts (column + line) from scratch; style single-type
  charts or finish in PowerPoint
- image formats are limited to BMP/GIF/JPEG/PNG/TIFF/WMF — webp assets must
  be converted before embedding
- generated decks still need a real PowerPoint review when fidelity or
  accessibility matters

## Agent-visible summary

Result context: artifacts (pptx/pdf/previews), backend per step, validation
checks + warnings, and whether human review is recommended.
