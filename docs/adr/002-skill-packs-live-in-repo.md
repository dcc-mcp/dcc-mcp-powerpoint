# ADR 002 — PowerPoint Skill Pack 随仓库发布

- **Status**: Accepted
- **Date**: 2026-08-16
- **Related**: dcc-mcp-maya skill layout, dcc-mcp-office ADR-006

## Context

Skill packs can ship in the adapter repo (maya style:
`src/dcc_mcp_maya/skills/...`) or in a separate `-skills` repo. PowerPoint
packs are tightly coupled to the adapter's capabilities.

## Decision

- PowerPoint-specific packs (`powerpoint-deck`, `powerpoint-review`)
  live in `src/dcc_mcp_powerpoint/skills/` and ship with the package.
- Office-generic packs stay in `dcc-mcp-office/skills/`.
- A separate `dcc-mcp-powerpoint-skills` repo is created only if pack
  volume/consumer base demands it.

## Consequences

- One release train per repo; skills and capabilities cannot drift apart.
- Cross-DCC packs (`dcc.review-deck-from-renders`) are authored here but
  reference Maya/Houdini/Blender adapters only through their public
  capabilities (no hard coupling).
