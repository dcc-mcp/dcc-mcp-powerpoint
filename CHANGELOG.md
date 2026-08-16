# Changelog

## [0.2.0] - 2026-08-16

First release of the PowerPoint adapter over the dcc-mcp-office core.

### Added

- Deck pipeline: IR → Open XML host compile → inspect (deck/slide/shape/notes
  structure), official brand identity, 16-slide showcase deck,
  capability-boundary showcase with patch variants.
- Smart layers API: list / assign / reorder / recolor / visibility.
- Deck patch engine: edit-deck operations with failing-operation index
  reporting.
- Plugin registry with a deck-stats example plugin.
- Skill packs: powerpoint-deck, powerpoint-edit, powerpoint-layers,
  powerpoint-plugins, powerpoint-review (official lint gate in CI).
- Standalone distribution via PyOxidizer (dcc-mcp-powerpoint.exe), built
  against the pinned dcc-office-host source.

### Changed

- Host client drives the office-rpc JSON-RPC surface over stdio.
- Confined-environment stdio fixes; per-script font slot; self-implemented
  issue analyzer per the dependency policy.

### Fixed

- Skill script contract compliance.
