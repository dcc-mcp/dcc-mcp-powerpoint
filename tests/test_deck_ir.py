"""Deck IR contract tests — the domain core of the adapter."""

from __future__ import annotations

from pathlib import Path

import pytest

from dcc_mcp_powerpoint.deck_ir import (
    IR_VERSION,
    IrValidationError,
    artifact_stem,
    load_deck_ir,
    parse_envelope,
)

EXAMPLES = Path(__file__).resolve().parent.parent / "examples"


def _valid_envelope() -> dict:
    return {
        "schema_version": IR_VERSION,
        "kind": "presentation",
        "document_id": "draft:t",
        "metadata": {"title": "T", "author": "a", "language": "zh-CN"},
        "document": {
            "slides": [
                {
                    "semantic_layout": "bullets",
                    "title": "S1",
                    "content_blocks": [{"type": "bullets", "items": ["a", "b"]}],
                }
            ]
        },
    }


def test_example_envelope_loads() -> None:
    env = load_deck_ir(EXAMPLES / "dcc_mcp_framework_intro.json")
    assert env.kind == "presentation"
    assert len(env.document.slides) == 16
    assert env.document.slides[0].semantic_layout == "title_cover"
    assert env.metadata.title == "DCC-MCP 框架介绍"


def test_wrong_schema_version_rejected() -> None:
    raw = _valid_envelope()
    raw["schema_version"] = "office-ir/9.9"
    with pytest.raises(IrValidationError, match="schema_version"):
        parse_envelope(raw)


def test_wrong_kind_rejected() -> None:
    raw = _valid_envelope()
    raw["kind"] = "workbook"
    with pytest.raises(IrValidationError, match="kind"):
        parse_envelope(raw)


def test_empty_slides_rejected() -> None:
    raw = _valid_envelope()
    raw["document"]["slides"] = []
    with pytest.raises(IrValidationError, match="slides"):
        parse_envelope(raw)


def test_unknown_block_type_rejected() -> None:
    raw = _valid_envelope()
    raw["document"]["slides"][0]["content_blocks"] = [{"type": "nope"}]
    with pytest.raises(IrValidationError, match="unknown content block type"):
        parse_envelope(raw)


def test_missing_input_file_rejected() -> None:
    with pytest.raises(IrValidationError, match="not found"):
        load_deck_ir(EXAMPLES / "does-not-exist.json")


def test_artifact_stem_sanitizes_document_ids() -> None:
    assert artifact_stem("draft:dcc-mcp-framework-intro") == "draft-dcc-mcp-framework-intro"
    assert artifact_stem("ppt:87f1") == "ppt-87f1"
    assert artifact_stem("///") == "deck"
