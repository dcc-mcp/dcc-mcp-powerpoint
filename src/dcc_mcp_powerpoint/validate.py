"""Structural validation for compiled decks (proposal §18.1).

Structural checks only — visual checks come from the rendered previews.
Honest reporting: every check returns pass/warn with a reason; the deck
contract itself is enforced by deck_ir at load time.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .deck_ir import DOCUMENT_KIND, IR_VERSION, DeckEnvelope

MAX_BULLETS_PER_SLIDE = 6
MAX_BULLET_CHARS = 80


@dataclass(frozen=True)
class Check:
    name: str
    ok: bool
    detail: str = ""


def validate_envelope(envelope: DeckEnvelope) -> dict[str, Any]:
    """Validate a DeckEnvelope; returns {ok, checks, warnings}."""
    checks: list[Check] = []
    warnings: list[str] = []

    checks.append(Check("schema_version", envelope.schema_version == IR_VERSION, envelope.schema_version))
    checks.append(Check("document_kind", envelope.kind == DOCUMENT_KIND, envelope.kind))
    checks.append(Check("has_slides", len(envelope.document.slides) > 0, f"{len(envelope.document.slides)} slides"))
    checks.append(Check("has_title", bool(envelope.metadata.title.strip()), envelope.metadata.title))

    slides = envelope.document.slides
    cover_layouts = {"title_cover", "section_cover", "closing"}
    for index, slide in enumerate(slides, start=1):
        if slide.semantic_layout not in cover_layouts and not slide.title:
            warnings.append(f"slide {index}: content slide without a title")
        for block in slide.content_blocks:
            if block["type"] == "bullets":
                items = block.get("items", [])
                if len(items) > MAX_BULLETS_PER_SLIDE:
                    warnings.append(f"slide {index}: {len(items)} bullets exceed {MAX_BULLETS_PER_SLIDE}")
                for item in items:
                    if len(item) > MAX_BULLET_CHARS:
                        warnings.append(f"slide {index}: bullet longer than {MAX_BULLET_CHARS} chars — overflow risk")
    ok = all(c.ok for c in checks)
    return {"ok": ok, "checks": [c.__dict__ for c in checks], "warnings": warnings}


def validate_artifacts(paths: list[str | Path]) -> dict[str, Any]:
    """Check produced artifacts exist and are non-empty (proposal §17)."""
    results = []
    for raw in paths:
        p = Path(raw)
        results.append({"path": str(p), "ok": p.is_file() and p.stat().st_size > 0})
    return {"ok": all(r["ok"] for r in results), "artifacts": results}
