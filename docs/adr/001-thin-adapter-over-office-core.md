# ADR 001 — PowerPoint 仓库只做应用语义

- **Status**: Accepted
- **Date**: 2026-08-16
- **Related**: dcc-mcp-office ADR-006, proposal §5

## Context

The platform proposal sketches one monolithic office repo. The ecosystem
pattern is thin per-DCC adapters over `dcc-mcp-core`; Office adds a second
shared layer (`dcc-mcp-office`) because 8 applications share one COM/Open
XML runtime.

## Decision

- `dcc-mcp-PowerPoint` owns only PowerPoint semantics: deck/slide tools,
  template registry entries, PowerPoint skill packs, host launcher.
- Protocol, IR envelope, C# runtime, Open XML worker, jobs and security
  policy come from `dcc-mcp-office` (published crates + release artifacts).
- No PowerPoint-specific COM code may drift into the shared runtime without
  an ADR there (keeps the shared core reusable by Word/Excel/SUA).

## Consequences

- This repo stays small and releases independently.
- PowerPoint fixes cannot destabilize Word/Excel — and vice versa.
- Cross-repo contract tests are required (dcc-mcp-office-testkit).
