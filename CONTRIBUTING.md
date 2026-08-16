# CONTRIBUTING — dcc-mcp-PowerPoint

Thin adapter over `dcc-mcp-office`: PowerPoint semantics only. The shared
engineering agreement (first principles, contract-first, SOLID, Clean
Architecture, no code smells) lives in
`dcc-mcp-office/CONTRIBUTING.md` and applies here unchanged.

Python-specific rules:

- Typed public API (`from __future__ import annotations`), dataclasses for
  configuration, no mutable defaults.
- `__init__.py` imports only the public surface — no side effects, no
  COM/process spawning at import time.
- The launcher never kills a user-started POWERPNT.EXE (proposal §8.3).
- Contract constants (pipe prefix, app name, protocol version) must match
  `dcc-mcp-office-protocol`; a test pins them (`tests/test_version.py`).

Gates:

```bash
ruff check src tests
pytest
python -m py_compile src/dcc_mcp_powerpoint
dcc-mcp-cli lint src/dcc_mcp_powerpoint/skills --warnings-as-errors --non-interactive
```
