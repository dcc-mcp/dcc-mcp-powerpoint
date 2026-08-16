# Learnings — powerpoint-deck

Dated lessons applied to this skill (learnings loop; append, never edit
history).

## 2026-08-16 — showcase deck optimization + ecosystem research

- Web research across document-pptx (vasilyu1983/AI-Agents-public),
  IBM/chuk-mcp-pptx, and Tencent/Youtu-agent ppt_gen confirmed the core
  design: template-first with layout resolution by name, schema-driven
  content (their YAML type_map + field constraints ≙ our Deck IR +
  validation), and a component registry with LLM-friendly schemas (≙
  office-tools registry).
- Adopted: alt text on every picture (a11y as authoring), inventory tool
  before repair, decision rules + known-limits in SKILL.md, and this
  learnings loop.
- Gotcha found live: a site asset named .png was actually WEBP content —
  python-pptx rejects WEBP; always verify magic bytes, not extensions.
- Showcase imagery now embeds normalized 480x270 cards (16:9, alpha-padded);
  image_grid layout renders slide.images as captioned cards with explicit
  missing_asset notes.
