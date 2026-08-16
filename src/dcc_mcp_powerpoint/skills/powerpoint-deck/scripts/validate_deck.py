"""powerpoint-deck / validate_deck — validate a Deck IR or artifact set."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

_THIS = Path(__file__).resolve()
try:
    from dcc_mcp_powerpoint.deck_ir import load_deck_ir
    from dcc_mcp_powerpoint.validate import validate_artifacts, validate_envelope
except ImportError:
    sys.path.insert(0, str(_THIS.parents[4]))
    from dcc_mcp_powerpoint.deck_ir import load_deck_ir
    from dcc_mcp_powerpoint.validate import validate_artifacts, validate_envelope


def run(params: dict) -> None:
    target = Path(params["input"])
    if target.is_dir():
        report = validate_artifacts(sorted(str(p) for p in target.rglob("*") if p.is_file()))
        payload = {"success": report["ok"], "message": "artifact validation", "context": report}
    else:
        envelope = load_deck_ir(target)
        report = validate_envelope(envelope)
        payload = {"success": report["ok"], "message": f"deck '{envelope.document_id}' validated", "context": report}
    print(json.dumps(payload, ensure_ascii=False))


def _force_utf8_stdio() -> None:
    """Deterministic output contract: stdout/stderr are always UTF-8.

    On Windows, a piped subprocess stdout defaults to the ANSI codepage
    (charmap) and fails on CJK text. The gateway reads JSON from stdout, so
    the encoding is part of the script contract.
    """
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
        parser = argparse.ArgumentParser(description="Validate a Deck IR or artifacts")
        parser.add_argument("--input", required=True)
        params = vars(parser.parse_args())
    try:
        run(params)
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"success": False, "message": str(exc), "context": {}}, ensure_ascii=False))
        sys.exit(1)


if __name__ == "__main__":
    main()
