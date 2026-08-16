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

## Dependency policy (2026-08-16)

- **Python runtime: stdlib only** (zipfile / xml.etree / json / dataclasses /
  subprocess / ctypes). No python-pptx, no Pillow, no pywin32 at runtime.
- Heavy interfaces live in **our own C# host** (`dcc-mcp-office`): the Open
  XML worker and the COM renderer. Python talks to it over the office-rpc
  contract (subprocess stdin/stdout JSON-RPC first, named pipe per proposal
  §12 later).
- python-pptx / Pillow / pywin32 are **dev/test-only** (test oracle,
  fixtures, experiments) — package modules never import them at runtime;
  current exceptions (analyze/inventory readers) migrate to the C# host.
- The C# host adds **no non-Microsoft NuGet dependencies**: OOXML via
  in-box System.IO.Packaging + LINQ to XML (net8.0-windows), COM via BCL
  interop.
- Third-party assets (templates, fonts, showcase imagery) enter the repo
  only after license verification and self-hosting — never fetched from
  external sources at build or runtime.

Gates:

```bash
ruff check src tests
pytest
python -m py_compile src/dcc_mcp_powerpoint
dcc-mcp-cli lint src/dcc_mcp_powerpoint/skills --warnings-as-errors --non-interactive
```
