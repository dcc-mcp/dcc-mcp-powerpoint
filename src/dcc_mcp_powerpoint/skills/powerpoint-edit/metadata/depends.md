# Dependencies — powerpoint-edit

This skill depends on the following skill packages
(`metadata.dcc-mcp.depends` in SKILL.md):

## powerpoint-deck

Used for the render → look → fix loop after a patch:

- `render_deck` — export PDF + per-slide PNG previews of the patched deck
  so the agent (and the human) can look at the result
- `analyze_deck` — path-addressed issues report with concrete fix hints
- `inventory_deck` — read-only inventory before touching decks the agent
  did not generate

The patch engine itself (ppt-patch/1.0, edits.py) is self-contained; the
dependency exists because an edit is only complete once the result has
been rendered and checked.
