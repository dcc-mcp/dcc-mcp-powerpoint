"""powerpoint-layers / layer_recolor — retint a layer (fill/line/text)

Parameter resolution order (dcc-mcp-core execute_script convention):
1. stdin JSON: {...}
2. CLI flags
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from pptx import Presentation

from dcc_mcp_powerpoint.layers import recolor_layer


def _force_utf8_stdio() -> None:
    """Deterministic output contract: stdout/stderr are always UTF-8."""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def _as_bool(value) -> bool:
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in ("1", "true", "yes", "on")


def _fail(message: str) -> None:
    print(json.dumps({"success": False, "message": message, "context": {}}, ensure_ascii=False))


def _load(params: dict):
    path = Path(params["input"])
    if not path.is_file():
        _fail("input not found: " + str(path))
        return None
    return Presentation(str(path)), path


def _save(prs, path: Path, params: dict) -> str:
    target = Path(params["output"]) if params.get("output") else path
    target.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(target))
    return str(target)


def _slides(params: dict):
    raw = params.get("slides")
    if raw is None:
        return None
    if not isinstance(raw, list) or not all(isinstance(i, int) for i in raw):
        raise ValueError("'slides' must be a list of integers")
    return raw

def run(params: dict) -> None:
    loaded = _load(params)
    if loaded is None:
        return
    prs, path = loaded
    report = recolor_layer(prs, str(params["layer"]), fill=params.get("fill"), line=params.get("line"), text=params.get("text"), slides=_slides(params))
    target = _save(prs, path, params)
    print(json.dumps({"success": True, "message": f"layer '{params['layer']}' recolored ({report['applied']} style applications)", "context": {"output": target, "recolor": report}}, ensure_ascii=False))


def main() -> None:
    _force_utf8_stdio()
    params = {}
    if not sys.stdin.isatty():
        raw = sys.stdin.read()
        if raw.strip():
            try:
                params = json.loads(raw)
            except json.JSONDecodeError:
                params = {}
    if not params:
        parser = argparse.ArgumentParser(description="retint a layer (fill/line/text)")
        parser.add_argument("--input", dest="input", required=True, help="PPTX file path")
        parser.add_argument("--layer", dest="layer", required=True, help="layer name")
        parser.add_argument("--fill", dest="fill", help="hex RRGGBB fill")
        parser.add_argument("--line", dest="line", help="hex RRGGBB line")
        parser.add_argument("--text", dest="text", help="hex RRGGBB text")
        parser.add_argument("--slides", dest="slides", help="optional comma-separated 1-based slide indexes")
        parser.add_argument("--output", dest="output", help="optional output path (default in place)")
        params = vars(parser.parse_args())
    try:
        run(params)
    except Exception as exc:  # noqa: BLE001 — surface as structured error
        _fail(str(exc))
        sys.exit(1)


if __name__ == "__main__":
    main()
