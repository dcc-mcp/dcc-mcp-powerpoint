"""powerpoint-layers / layer_assign — tag existing shapes into a layer

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

from dcc_mcp_powerpoint.layers import assign_layer


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
    select = params.get("select") or {}
    slides = _slides({"slides": select.get("slides")})
    report = assign_layer(prs, str(params["layer"]), slides=slides, shape=select.get("shape"), shape_id=select.get("id"), all_shapes=bool(select.get("all_shapes", False)))
    target = _save(prs, path, params)
    print(json.dumps({"success": True, "message": f"{report['tagged']} shapes tagged into layer '{params['layer']}'", "context": {"output": target, "assign": report}}, ensure_ascii=False))


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
        parser = argparse.ArgumentParser(description="tag existing shapes into a layer")
        parser.add_argument("--input", dest="input", required=True, help="PPTX file path")
        parser.add_argument("--layer", dest="layer", required=True, help="target layer name")
        parser.add_argument("--select", dest="select", help="JSON selector object (stdin only)")
        params = vars(parser.parse_args())
    try:
        run(params)
    except Exception as exc:  # noqa: BLE001 — surface as structured error
        _fail(str(exc))
        sys.exit(1)


if __name__ == "__main__":
    main()
