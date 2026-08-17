# Changelog

## [0.2.2](https://github.com/dcc-mcp/dcc-mcp-powerpoint/compare/v0.2.1...v0.2.2) (2026-08-17)


### Bug Fixes

* upload only the two standalone artifacts (exe + bundled host), not the whole dist tree ([a946a24](https://github.com/dcc-mcp/dcc-mcp-powerpoint/commit/a946a24f9891b5eb3bc5eb176dc14f68d687ed0d))

## [0.2.1](https://github.com/dcc-mcp/dcc-mcp-powerpoint/compare/v0.2.0...v0.2.1) (2026-08-17)


### Features

* integrate dcc-office-host v0.2.0 capability surface ([#15](https://github.com/dcc-mcp/dcc-mcp-powerpoint/issues/15)) ([c2d06d8](https://github.com/dcc-mcp/dcc-mcp-powerpoint/commit/c2d06d8ec9976f02d05eabfe1751f6a3f1365b45))


### Bug Fixes

* correct workflow expression syntax (literal backslashes broke the token reference) ([77adc36](https://github.com/dcc-mcp/dcc-mcp-powerpoint/commit/77adc3631e41d079a796ea9ec368c58eda6e2f6d))

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
