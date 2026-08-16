"""Structural validation tests."""

from __future__ import annotations

from pathlib import Path

from dcc_mcp_powerpoint.deck_ir import load_deck_ir
from dcc_mcp_powerpoint.validate import validate_artifacts, validate_envelope

EXAMPLES = Path(__file__).resolve().parent.parent / "examples"


def test_example_deck_passes_structural_checks() -> None:
    envelope = load_deck_ir(EXAMPLES / "dcc_mcp_framework_intro.json")
    report = validate_envelope(envelope)
    assert report["ok"], report["checks"]
    assert report["warnings"] == []


def test_artifact_validation(tmp_path: Path) -> None:
    good = tmp_path / "good.txt"
    good.write_text("x", encoding="utf-8")
    empty = tmp_path / "empty.txt"
    empty.write_text("", encoding="utf-8")
    missing = tmp_path / "nope.txt"
    report = validate_artifacts([str(good), str(empty), str(missing)])
    assert not report["ok"]
    assert [r["ok"] for r in report["artifacts"]] == [True, False, False]
