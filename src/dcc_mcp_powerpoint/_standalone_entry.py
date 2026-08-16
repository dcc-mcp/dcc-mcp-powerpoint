"""Dual-purpose entry point for the PyOxidizer standalone executable.

- CLI mode: dcc-mcp-powerpoint <command> (compile/inspect/validate/render/
  version) — compile and inspect run through the bundled dcc-office-host
  (stdlib-only client); render uses the desktop COM backend.
- Skill-script mode: dcc-mcp-powerpoint scripts/generate_deck.py ... runs
  the script with the embedded interpreter (gateway execute_script parity),
  mirroring the dcc-mcp-photoshop standalone pattern.
"""

from __future__ import annotations

import argparse
import json
import os
import runpy
import sys
from collections.abc import Sequence
from pathlib import Path

from .host_client import compile_deck, inspect_deck
from .render import render_deck
from .validate import validate_artifacts

_PYTHON_SCRIPT_SUFFIXES = frozenset({".py", ".pyw"})


def _is_skill_script_invocation(argv: Sequence[str]) -> bool:
    if len(argv) < 2:
        return False
    script = Path(argv[1])
    return script.suffix.lower() in _PYTHON_SCRIPT_SUFFIXES and script.is_file()


def _run_skill_script(argv: Sequence[str]) -> None:
    script = str(Path(argv[1]).resolve())
    original_argv = sys.argv
    sys.argv = [script, *argv[2:]]
    try:
        runpy.run_path(script, run_name="__main__")
    finally:
        sys.argv = original_argv


def _cmd_compile(args: argparse.Namespace) -> int:
    result = compile_deck(args.input, args.output)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("success") else 1


def _cmd_inspect(args: argparse.Namespace) -> int:
    result = inspect_deck(args.input)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("success") else 1


def _cmd_validate(args: argparse.Namespace) -> int:
    report = validate_artifacts(sorted(str(p) for p in Path(args.input).rglob("*") if p.is_file())) if Path(args.input).is_dir() else validate_artifacts([args.input])
    print(json.dumps({"success": report["ok"], "context": report}, ensure_ascii=False))
    return 0 if report["ok"] else 1


def _cmd_render(args: argparse.Namespace) -> int:
    report = render_deck(args.input, args.output, pdf=not args.no_pdf, previews=not args.no_previews)
    print(json.dumps(report, ensure_ascii=False))
    return 0 if report.get("success") else 1


def main(argv: Sequence[str] | None = None) -> None:
    resolved_argv = list(sys.argv if argv is None else argv)
    os.environ.setdefault("DCC_MCP_PYTHON_EXECUTABLE", sys.executable)
    if _is_skill_script_invocation(resolved_argv):
        _run_skill_script(resolved_argv)
        return

    parser = argparse.ArgumentParser(prog="dcc-mcp-powerpoint", description="Standalone PowerPoint adapter (bundled dcc-office-host)")
    sub = parser.add_subparsers(dest="command", required=True)
    c = sub.add_parser("compile", help="Deck IR JSON -> PPTX via the bundled host")
    c.add_argument("--input", required=True)
    c.add_argument("--output", required=True)
    c.set_defaults(func=_cmd_compile)
    i = sub.add_parser("inspect", help="inspect a PPTX (path addressing)")
    i.add_argument("--input", required=True)
    i.set_defaults(func=_cmd_inspect)
    v = sub.add_parser("validate", help="validate artifacts exist and are non-empty")
    v.add_argument("--input", required=True)
    v.set_defaults(func=_cmd_validate)
    r = sub.add_parser("render", help="render PPTX -> PDF + previews via desktop COM")
    r.add_argument("--input", required=True)
    r.add_argument("--output", required=True)
    r.add_argument("--no-pdf", action="store_true")
    r.add_argument("--no-previews", action="store_true")
    r.set_defaults(func=_cmd_render)
    parser.add_argument("--version", action="version", version="dcc-mcp-powerpoint 0.1.0")
    args = parser.parse_args(resolved_argv[1:])
    raise SystemExit(args.func(args))


if __name__ == "__main__":
    main()
