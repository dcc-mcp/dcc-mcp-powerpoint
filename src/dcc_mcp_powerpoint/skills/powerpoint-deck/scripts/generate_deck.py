"""powerpoint-deck / generate_deck — Deck IR → PPTX (+ PDF/previews).

Parameter resolution order (dcc-mcp-core execute_script convention):
1. stdin JSON: {"input": ..., "output_dir": ..., "render": ..., "previews": ...}
2. CLI flags: --input --out --render/--no-render --previews/--no-previews
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

_THIS = Path(__file__).resolve()
try:
    from dcc_mcp_powerpoint.compiler import compile_deck
    from dcc_mcp_powerpoint.deck_ir import artifact_stem, load_deck_ir
    from dcc_mcp_powerpoint.render import render_deck
    from dcc_mcp_powerpoint.validate import validate_artifacts, validate_envelope
except ImportError:  # running from a source checkout without install
    sys.path.insert(0, str(_THIS.parents[4]))
    from dcc_mcp_powerpoint.compiler import compile_deck
    from dcc_mcp_powerpoint.deck_ir import artifact_stem, load_deck_ir
    from dcc_mcp_powerpoint.render import render_deck
    from dcc_mcp_powerpoint.validate import validate_artifacts, validate_envelope


def run(params: dict) -> None:
    envelope = load_deck_ir(params["input"])
    report = validate_envelope(envelope)
    if not report["ok"]:
        print(json.dumps({"success": False, "message": "deck IR failed structural validation", "context": report}, ensure_ascii=False))
        return

    out_dir = Path(params.get("output_dir", "output"))
    pptx = compile_deck(envelope, out_dir / f"{artifact_stem(envelope.document_id)}.pptx")
    artifacts = [str(pptx)]

    backend = "openxml"
    render_report: dict = {"backend": None, "reason": "render skipped by request"}
    if params.get("render", True):
        render_report = render_deck(
            pptx,
            out_dir,
            pdf=envelope.document.export_policy.pdf,
            previews=params.get("previews", envelope.document.export_policy.slide_previews),
        )
        if render_report.get("backend"):
            backend = render_report["backend"]
        if render_report.get("pdf"):
            artifacts.append(render_report["pdf"])
        if render_report.get("previews"):
            artifacts.extend(render_report["previews"])

    artifact_report = validate_artifacts(artifacts)
    print(
        json.dumps(
            {
                "success": artifact_report["ok"],
                "message": f"deck '{envelope.document_id}' compiled ({len(envelope.document.slides)} slides)",
                "context": {
                    "document_id": envelope.document_id,
                    "artifacts": artifacts,
                    "backend": backend,
                    "render": render_report,
                    "validation": report,
                    "artifacts_ok": artifact_report,
                },
            },
            ensure_ascii=False,
        )
    )


def main() -> None:
    params: dict = {}
    if not sys.stdin.isatty():
        raw = sys.stdin.read()
        if raw.strip():
            try:
                params = json.loads(raw)
            except json.JSONDecodeError:
                params = {}
    if not params:
        parser = argparse.ArgumentParser(description="Deck IR → PPTX (+PDF/previews)")
        parser.add_argument("--input", required=True, help="Deck IR JSON path")
        parser.add_argument("--out", dest="output_dir", default="output", help="output directory")
        parser.add_argument("--render", dest="render", action=argparse.BooleanOptionalAction, default=True)
        parser.add_argument("--previews", dest="previews", action=argparse.BooleanOptionalAction, default=True)
        params = vars(parser.parse_args())
    try:
        run(params)
    except Exception as exc:  # noqa: BLE001 — surface as structured error
        print(json.dumps({"success": False, "message": str(exc), "context": {}}, ensure_ascii=False))
        sys.exit(1)


if __name__ == "__main__":
    main()
