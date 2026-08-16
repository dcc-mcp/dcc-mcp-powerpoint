"""powerpoint-edit / edit_deck — apply a ppt-patch/1.0 document to a PPTX.

Parameter resolution order (dcc-mcp-core execute_script convention):
1. stdin JSON: {"input": ..., "patch": ..., "output": ...}
2. CLI flags: --input --patch --output
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from dcc_mcp_powerpoint.edits import PatchApplyError, PatchValidationError, apply_patch


def _force_utf8_stdio() -> None:
    """Deterministic output contract: stdout/stderr are always UTF-8."""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def run(params: dict) -> None:
    patch_path = Path(params["patch"])
    if not patch_path.is_file():
        print(json.dumps({"success": False, "message": f"patch not found: {patch_path}", "context": {}}, ensure_ascii=False))
        return
    try:
        patch = json.loads(patch_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        print(json.dumps({"success": False, "message": f"patch is not valid JSON: {exc}", "context": {}}, ensure_ascii=False))
        return
    try:
        report = apply_patch(params["input"], patch, params.get("output"))
    except (PatchValidationError, PatchApplyError) as exc:
        print(json.dumps({"success": False, "message": str(exc), "context": {"input": params["input"]}}, ensure_ascii=False))
        sys.exit(1)
    print(
        json.dumps(
            {
                "success": True,
                "message": f"applied {len(report['operations'])} operations -> {report['output']} ({report['slides_after']} slides)",
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
        parser = argparse.ArgumentParser(description="Apply a ppt-patch/1.0 document to a PPTX")
        parser.add_argument("--input", required=True, help="PPTX file to modify")
        parser.add_argument("--patch", required=True, help="ppt-patch/1.0 JSON document")
        parser.add_argument("--output", help="output path (default <stem>-patched.pptx)")
        params = vars(parser.parse_args())
    try:
        run(params)
    except Exception as exc:  # noqa: BLE001 — surface as structured error
        print(json.dumps({"success": False, "message": str(exc), "context": {}}, ensure_ascii=False))
        sys.exit(1)


if __name__ == "__main__":
    main()
