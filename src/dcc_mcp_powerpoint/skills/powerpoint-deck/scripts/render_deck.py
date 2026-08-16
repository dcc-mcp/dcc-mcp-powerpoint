"""powerpoint-deck / render_deck — PPTX → PDF + previews via desktop COM."""

from __future__ import annotations

import argparse
import json
import sys

from dcc_mcp_powerpoint.render import render_deck


def _force_utf8_stdio() -> None:
    """Deterministic output contract: stdout/stderr are always UTF-8."""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def run(params: dict) -> None:
    report = render_deck(
        params["input"],
        params.get("output_dir", "output"),
        pdf=params.get("pdf", True),
        previews=params.get("previews", True),
    )
    print(json.dumps({"success": report.get("success", False), "message": "deck render", "context": report}, ensure_ascii=False))


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
        parser = argparse.ArgumentParser(description="Render PPTX to PDF + previews via COM")
        parser.add_argument("--input", required=True)
        parser.add_argument("--out", dest="output_dir", default="output")
        parser.add_argument("--pdf", dest="pdf", action=argparse.BooleanOptionalAction, default=True)
        parser.add_argument("--previews", dest="previews", action=argparse.BooleanOptionalAction, default=True)
        params = vars(parser.parse_args())
    try:
        run(params)
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"success": False, "message": str(exc), "context": {}}, ensure_ascii=False))
        sys.exit(1)


if __name__ == "__main__":
    main()
