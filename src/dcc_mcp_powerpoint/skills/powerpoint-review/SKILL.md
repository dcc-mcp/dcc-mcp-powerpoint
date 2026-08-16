---
name: powerpoint-review
description: >-
  Build a production review deck from DCC renders (Maya/Houdini/Blender/Unreal)
  plus shot/asset metadata: title page, shot list, per-shot slides with
  version/artist/date and notes, PDF + per-page previews, and backlinks to the
  source assets. Use for dailies and asset review decks.
license: MIT
allowed-tools: Bash Read
metadata:
  dcc-mcp:
    dcc: powerpoint
    layer: domain
    stage: review
    version: 0.1.1
    tags:
      - powerpoint
      - review
      - dcc
      - render
      - deck
    search-hint: >-
      review deck, dailies deck, shot review, asset review, render comparison,
      review-deck-from-renders
    tools: tools.yaml
---

# powerpoint-review (Review stage)

`dcc.review-deck-from-renders` (proposal §15.7): production review decks
straight from DCC pipelines. Deck assembly reuses the `powerpoint-deck`
pipeline (Deck IR → Open XML compile → COM render → validation).

## Input contract

- `input` — path to a shots manifest JSON:
  `{"title": ..., "shots": [{"name", "version", "artist", "date", "image", "notes"}]}`

## Scripts

- `review_deck_from_renders` — manifest → Deck IR → PPTX (+ PDF/previews)

## Validation rules

- every shot becomes a slide; a missing render gets a `missing_asset` note,
  never a silently dropped shot
- artifacts must exist and be non-empty

## Agent-visible summary

Deck path, PDF path, previews, per-shot coverage list, and any missing or
blocked assets requiring human follow-up.
