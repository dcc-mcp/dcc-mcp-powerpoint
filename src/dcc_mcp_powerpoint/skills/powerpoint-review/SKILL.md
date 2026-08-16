---
name: powerpoint-review
description: "dcc.review-deck-from-renders — build a production review deck from DCC renders (Maya/Houdini/Blender/Unreal) + shot/asset metadata: title page, shot/asset grouping, version comparisons, notes pages, PDF + per-page previews, backlinks to source assets."
dcc: powerpoint
version: "0.1.0"
license: "MIT"
compatibility: "Windows, Office 2019+ / Microsoft 365; inputs from any DCC-MCP render pipeline"
tags: ["powerpoint", "review", "dcc", "render", "deck"]
capabilities:
  - powerpoint.deck.generate
  - powerpoint.slide.compose
  - office.batch.convert
---

# powerpoint-review (dcc.review-deck-from-renders)

Production review decks straight from DCC pipelines (proposal §15.7).

## Input contract

- renders: images/videos from Maya/Houdini/Blender/Unreal
- shot/asset metadata: names, versions, artists, dates, performance data
- review notes (optional)

## Planning steps

1. Group by shot/asset; pair old vs new versions for comparisons.
2. Map groups to semantic layouts (full_bleed_image, comparison,
   kpi_dashboard for performance data).
3. Compose Deck IR with backlinks to source assets.
4. Generate via `powerpoint.deck.generate`; export PDF + per-page previews.

## Provider choice

Desktop COM for final render/export; Open XML for base construction.

## Safety confirmation

Read-only review generation → no confirmation; publishing to shared
locations → confirm.

## Validation rules

Every slide must carry version info; comparisons must show both versions;
previews must be checked for overflow before delivery.

## Failure compensation

Missing renders → placeholder slide with an explicit `missing_asset` note in
the report; never silently drop a shot from the review.

## Artifact naming

`review-<project>-<date>-v<n>.{pptx,pdf}` + per-page PNGs; artifact records
carry `source_document_id` backlinks to the DCC assets.

## Agent-visible summary

Deck path, PDF path, previews, per-shot coverage list, and any missing or
blocked assets requiring human follow-up.
