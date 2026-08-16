---
name: powerpoint-edit
description: >-
  Modify an existing PPTX with a structured patch document (ppt-patch/1.0):
  set text, speaker notes, replace pictures, recolor/outline shapes,
  show/hide and delete shapes, add/delete/move slides (new slides compile
  from Deck IR layouts), and tag shapes into semantic layers. Selectors
  address shapes by slide, layer, name, id or position — never by raw
  coordinates. All-or-nothing: a failed operation aborts before anything is
  saved. Use whenever an agent must change an existing deck instead of
  regenerating it.
license: MIT
allowed-tools: Bash Read
metadata:
  dcc-mcp:
    dcc: powerpoint
    layer: domain
    stage: authoring
    version: 0.1.0
    depends:
      - powerpoint-deck
    tags:
      - powerpoint
      - edit
      - modify
      - patch
      - pptx
    search-hint: >-
      edit pptx, modify powerpoint, change slide text, replace image,
      add slide, delete slide, move slide, speaker notes, patch deck
    tools: tools.yaml
---

# powerpoint-edit (Authoring stage)

Structured, coordinate-free deck modification. The agent writes a patch
document (ppt-patch/1.0) and applies it in one call; the engine resolves
selectors, executes operations in order, and saves only if every operation
succeeded — the input file is never left half-edited.

## Workflow

1. `slide_tree` (or `powerpoint-deck`'s `inventory_deck`) — read the
   shape tree: slide indexes, shape names/ids, layer tags, text previews
2. author the patch: one operation per change, ordered (add_slide before
   set_text that targets it, for example)
3. `edit_deck` applies it; a failed op reports its index and message, and
   the input stays untouched
4. finish with `powerpoint-deck`'s `render_deck` + `analyze_deck` to
   look at and check the result (render → look → fix loop)

## Patch contract (ppt-patch/1.0)

```json
{
  "schema_version": "ppt-patch/1.0",
  "metadata": {"title": "optional deck title for new-slide footers"},
  "operations": [
    {"op": "set_text", "select": {"slide": 3, "layer": "header"}, "text": "新标题"},
    {"op": "add_slide", "slide_ir": {"semantic_layout": "bullets", "title": "…",
      "content_blocks": [{"type": "bullets", "items": ["…"]}]}, "position": "end"}
  ]
}
```

### Selectors

All keys optional, combined with AND; at least one required:

| Key | Meaning |
|---|---|
| slide / slides | 1-based slide index, list of indexes, or "all" |
| layer | semantic layer tag (see powerpoint-layers) |
| shape | exact shape name or substring |
| id | shape cNvPr id |
| index | 1-based shape position on its slide |

### Operations

| Op | Fields | Effect |
|---|---|---|
| set_text | select, text | replace shape text |
| set_notes | slide, text | replace speaker notes |
| set_image | select, resource, alt? | replace a picture in place |
| set_shape_fill | select, color | solid fill RRGGBB |
| set_shape_line | select, color | solid outline RRGGBB |
| set_shape_visible | select, visible | native hidden attribute |
| delete_shape | select | remove shapes |
| add_slide | slide_ir, position? | compile a Deck IR slide (all layouts) |
| delete_slide | slide | remove a slide |
| move_slide | slide, to | reorder; to = index / "start" / "end" |
| assign_layer | layer, select | tag shapes into a layer |

## Decision rules

- edit before regenerate: patch what exists; regenerate only when the change
  is a rewrite
- keep edits semantic: select by layer or name, not by guessed z-position
- one concern per patch so failures are attributable; never batch unrelated
  edits
- new slides use the Deck IR semantic layouts — no raw coordinates

## Known limits

- text replacement keeps the first paragraph's formatting; per-run styling
  needs a later recolor pass (powerpoint-layers)
- deleted slides leave their part in the package (orphaned, harmless);
  PowerPoint cleans them on save
- charts/tables are edited as whole objects, not cell-by-cell

## Agent-visible summary

Per-operation ok/affected counts, output path, and the slide count after
the patch.
