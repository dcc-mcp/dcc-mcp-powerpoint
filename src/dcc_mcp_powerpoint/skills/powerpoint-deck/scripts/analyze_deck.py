"""powerpoint-deck / analyze_deck — issue analysis with path addressing.

Read-only. Self-implemented analyzer (OfficeCLI research): structure
(overflow, slide bounds), format (brand fonts, contrast), content (alt
text, notes). Every issue carries a path (/slide[i]/shape[j]) and a
concrete fix hint; undecidable checks are skipped, never guessed.
"""

from __future__ import annotations

import argparse
import json
import sys

from dcc_mcp_powerpoint.analyze import analyze_deck


def _force_utf8_stdio() -> None:
    """Deterministic output contract: stdout/stderr are always UTF-8."""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


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
        parser = argparse.ArgumentParser(description="Analyze a PPTX for issues (read-only)")
        parser.add_argument("--input", required=True)
        params = vars(parser.parse_args())
    try:
        report = analyze_deck(params["input"])
        report.setdefault("success", False)
        print(json.dumps({"success": report.pop("success", False), "message": f"analyzed: {report.get('count', 0)} issues", "context": report}, ensure_ascii=False))
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"success": False, "message": str(exc), "context": {}}, ensure_ascii=False))
        sys.exit(1)


if __name__ == "__main__":
    main()
