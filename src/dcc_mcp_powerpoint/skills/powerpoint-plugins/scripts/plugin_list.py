"""powerpoint-plugins / plugin_list — discover and validate installed plugins.

Parameter resolution order (dcc-mcp-core execute_script convention):
1. stdin JSON: {"paths": ...}
2. CLI flags: --paths
"""

from __future__ import annotations

import argparse
import json
import sys

from dcc_mcp_powerpoint.plugins import discover_plugins


def _force_utf8_stdio() -> None:
    """Deterministic output contract: stdout/stderr are always UTF-8."""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def run(params: dict) -> None:
    report = discover_plugins(params.get("paths"))
    plugins = report["plugins"]
    message = f"{len(plugins)} plugin(s) discovered"
    if report["errors"]:
        message += f", {len(report['errors'])} error(s)"
    print(json.dumps({"success": True, "message": message, "context": report}, ensure_ascii=False))


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
        parser = argparse.ArgumentParser(description="Discover PowerPoint plugins")
        parser.add_argument("--paths", help="optional plugin roots (os.pathsep-separated)")
        params = vars(parser.parse_args())
    try:
        run(params)
    except Exception as exc:  # noqa: BLE001 — surface as structured error
        print(json.dumps({"success": False, "message": str(exc), "context": {}}, ensure_ascii=False))
        sys.exit(1)


if __name__ == "__main__":
    main()
