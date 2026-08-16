"""powerpoint-edit / slide_tree — dump a shape tree for patch authoring.

Parameter resolution order (dcc-mcp-core execute_script convention):
1. stdin JSON: {"input": ..., "slide": ...}
2. CLI flags: --input --slide
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from pptx import Presentation

from dcc_mcp_powerpoint.layers import shape_tree


def _force_utf8_stdio() -> None:
    """Deterministic output contract: stdout/stderr are always UTF-8."""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def run(params: dict) -> None:
    path = Path(params["input"])
    if not path.is_file():
        print(json.dumps({"success": False, "message": f"input not found: {path}", "context": {}}, ensure_ascii=False))
        return
    slide = params.get("slide")
    if slide is not None:
        slide = int(slide)
    prs = Presentation(str(path))
    report = shape_tree(prs, str(path), slide=slide)
    print(
        json.dumps(
            {
                "success": True,
                "message": f"shape tree for {len(report['slides'])} slide(s) of {report['slide_count']}",
                "context": report,
            },
            ensure_ascii=False,
        )
    )


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
        parser = argparse.ArgumentParser(description="Dump the shape tree of a PPTX")
        parser.add_argument("--input", required=True, help="PPTX file path")
        parser.add_argument("--slide", type=int, help="optional 1-based slide index")
        params = vars(parser.parse_args())
    try:
        run(params)
    except Exception as exc:  # noqa: BLE001 — surface as structured error
        print(json.dumps({"success": False, "message": str(exc), "context": {}}, ensure_ascii=False))
        sys.exit(1)


if __name__ == "__main__":
    main()
