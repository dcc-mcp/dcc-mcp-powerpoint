"""Open XML compiler tests — Office-free, deterministic geometry."""

from __future__ import annotations

from pathlib import Path

import pytest
from pptx import Presentation

from dcc_mcp_powerpoint.compiler import LAYOUTS, compile_deck
from dcc_mcp_powerpoint.deck_ir import (
    DeckEnvelope,
    Metadata,
    PresentationIr,
    Slide,
    load_deck_ir,
)

EXAMPLES = Path(__file__).resolve().parent.parent / "examples"


def test_layout_registry_covers_example_deck() -> None:
    envelope = load_deck_ir(EXAMPLES / "dcc_mcp_framework_intro.json")
    used = {s.semantic_layout for s in envelope.document.slides}
    missing = used - set(LAYOUTS)
    assert not missing, f"layouts missing from registry: {missing}"


def test_compile_example_deck(tmp_path: Path) -> None:
    envelope = load_deck_ir(EXAMPLES / "dcc_mcp_framework_intro.json")
    out = compile_deck(envelope, tmp_path / "deck.pptx")
    assert out.is_file() and out.stat().st_size > 10_000
    reopened = Presentation(str(out))
    assert len(reopened.slides) == len(envelope.document.slides)
    # Speaker notes survive the compile step.
    first = reopened.slides[0]
    assert first.has_notes_slide
    assert "开场" in first.notes_slide.notes_text_frame.text


def test_compile_rejects_unknown_layout(tmp_path: Path) -> None:
    envelope = DeckEnvelope(
        schema_version="office-ir/1.0",
        kind="presentation",
        document_id="draft:bad",
        metadata=Metadata(title="bad"),
        document=PresentationIr(slides=(Slide(semantic_layout="does_not_exist"),)),
    )
    with pytest.raises(KeyError, match="does_not_exist"):
        compile_deck(envelope, tmp_path / "bad.pptx")


def test_all_example_slides_compile(tmp_path: Path) -> None:
    """Every layout in the registry compiles to a slide without error."""
    envelope = DeckEnvelope(
        schema_version="office-ir/1.0",
        kind="presentation",
        document_id="draft:layouts",
        metadata=Metadata(title="layouts"),
        document=PresentationIr(
            slides=tuple(
                Slide(
                    semantic_layout=layout,
                    title=f"L {layout}",
                    content_blocks=({"type": "bullets", "items": [f"{layout} item {i}" for i in range(3)]},),
                )
                for layout in sorted(LAYOUTS)
            )
        ),
    )
    out = compile_deck(envelope, tmp_path / "layouts.pptx")
    reopened = Presentation(str(out))
    assert len(reopened.slides) == len(LAYOUTS)
