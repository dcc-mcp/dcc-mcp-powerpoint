"""powerpoint-plugins / plugin_run — run one installed plugin.

Parameter resolution order (dcc-mcp-core execute_script convention):
1. stdin JSON: {"name": ..., "params": ..., "context": ..., "pptx": ..., "paths": ...}
2. CLI flags: --name --paths
"""

from __future__ import annotations

import argparse
import json
import sys

from dcc_mcp_powerpoint.plugins import PluginValidationError, run_plugin


def _force_utf8_stdio() -> None:
    """Deterministic output contract: stdout/stderr are always UTF-8."""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def run(params: dict) -> None:
    context = dict(params.get("context") or {})
    if params.get("pptx"):
        context.setdefault("pptx", params["pptx"])
    try:
        result = run_plugin(
            str(params["name"]),
            params.get("params"),
            context,
            paths=params.get("paths"),
        )
    except PluginValidationError as exc:
        print(json.dumps({"success": False, "message": str(exc), "context": {}}, ensure_ascii=False))
        sys.exit(1)
    success = bool(result.get("success"))
    print(
        json.dumps(
            {
                "success": success,
                "message": f"plugin '{result.get('plugin', params['name'])}' {'succeeded' if success else 'failed'}",
                "context": result,
            },
            ensure_ascii=False,
        )
    )
    if not success:
        sys.exit(1)


def main() -> None:
    _force_utf8_stdio()
    params: dict = {}
    if not sys.stdin.isatty():
        raw = sys.stdin.read()
        if raw.strip():
            try:
                params = json.loads(raw)
            except json.JSONDecodeError:
                params = {}
    if not params:
        parser = argparse.ArgumentParser(description="Run a PowerPoint plugin")
        parser.add_argument("--name", required=True, help="plugin name or plugin directory")
        parser.add_argument("--paths", help="optional plugin roots (os.pathsep-separated)")
        params = vars(parser.parse_args())
    try:
        run(params)
    except Exception as exc:  # noqa: BLE001 — surface as structured error
        print(json.dumps({"success": False, "message": str(exc), "context": {}}, ensure_ascii=False))
        sys.exit(1)


if __name__ == "__main__":
    main()
