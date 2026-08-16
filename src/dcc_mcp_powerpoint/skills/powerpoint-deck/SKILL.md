---
name: powerpoint-deck
description: "Generate an editable PowerPoint deck from structured content: requirement → outline → template + semantic layouts → Deck IR → Open XML compile → COM finalize → slide previews → structural + visual validation loop → PPTX/PDF/previews."
dcc: powerpoint
version: "0.1.0"
license: "MIT"
compatibility: "Windows, Office 2019+ / Microsoft 365; runtime from dcc-mcp-office"
tags: ["powerpoint", "deck", "generate", "template", "validation"]
capabilities:
  - powerpoint.deck.generate
  - powerpoint.slide.compose
  - powerpoint.slide.render
  - office.document.validate
---

# powerpoint-deck

Generate a deck from structured content (proposal §15.3/§15.4). Template-first:
the agent picks semantic layouts, never raw coordinates.

## Input contract

- structured content: per-slide intent + semantic layout + content blocks
- template: `brand://` URI + pinned version (registry in
  dcc-mcp-office/templates/)
- output targets: pptx, pdf, slide-previews
- validation rules: e.g. no_text_overflow, no_out_of_bounds, no_missing_fonts

## Planning steps

1. Content planner: outline + per-slide intent.
2. Template & layout resolver: choose `brand://` template and semantic
   layouts (title_cover, kpi_dashboard, technical_architecture, ...).
3. Build Deck IR (dcc-mcp-office-ir presentation schema).
4. Open XML compiler builds the base PPTX (fast, Office-free).
5. COM finalizer opens and completes native objects / layout / animation.
6. Render per-slide PNGs; run structural + visual validation.
7. Fail → generate patch → re-render; pass → publish artifacts.

## Provider choice

Open XML for base construction; desktop COM for finalize/render/export;
Graph only for OneDrive/SharePoint output targets.

## Safety confirmation

Generation writes only new files → no confirmation. Overwriting an existing
path follows the checkpoint + confirm policy from dcc-mcp-office-security.

## Validation rules

Structural: shape bounds, text overflow, missing placeholders, unresolved
media, occlusion candidates. Visual: per-slide preview review (vision model
where enabled). Both required — neither replaces the other.

## Failure compensation

Validation failure → patch loop (bounded retries) → if still failing, deliver
with a marked `needs_human_review` report; never silently downgrade quality.

## Artifact naming

`<workspace>/<deck-slug>-v<version>.{pptx,pdf}` + `previews/<deck-slug>/slide-<n>.png`.

## Agent-visible summary

Return: what was generated, template used, per-slide validation results,
preview paths, PDF path, and whether human review is recommended.
