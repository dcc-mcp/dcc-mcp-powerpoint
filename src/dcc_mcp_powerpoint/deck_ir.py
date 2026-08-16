"""Deck IR contract — mirrors the dcc-mcp-office-ir presentation schema 1:1.

Contract-first: this module is the *domain core* of the PowerPoint adapter.
The compiler (Open XML implementation) and the renderer (COM implementation)
both consume this contract, never raw coordinates.

JSON shape (snake_case, office-ir/1.0), see the Rust crate
dcc-mcp-office-ir for the authoritative schema:

    {
      "schema_version": "office-ir/1.0",
      "kind": "presentation",
      "document_id": "draft:review-deck",
      "metadata": {"title": "...", "author": "...", "language": "zh-CN"},
      "template": {"uri": "brand://...", "version": "1.0.0"},
      "resources": [],
      "document": {
        "slides": [
          {
            "semantic_layout": "bullets",
            "title": "...",
            "content_blocks": [
              {"type": "bullets", "items": ["..."]},
              {"type": "table", "header": true, "rows": [["a","b"]]}
            ],
            "speaker_notes": "..."
          }
        ],
        "export_policy": {"pdf": true, "slide_previews": true}
      },
      "validation": ["no_text_overflow", "no_out_of_bounds"],
      "outputs": ["pptx", "pdf", "slide-previews"]
    }
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, TypeAlias

IR_VERSION: str = "office-ir/1.0"
DOCUMENT_KIND: str = "presentation"


class IrValidationError(ValueError):
    """A Deck IR document violated the contract. Carries a json-path hint."""

    def __init__(self, path: str, message: str) -> None:
        super().__init__(f"[{path}] {message}")
        self.path = path
        self.message = message


@dataclass(frozen=True)
class TemplateRef:
    uri: str
    version: str


@dataclass(frozen=True)
class Resource:
    id: str
    uri: str
    mime: str | None = None


@dataclass(frozen=True)
class Metadata:
    title: str
    author: str = ""
    language: str = "zh-CN"


@dataclass(frozen=True)
class ExportPolicy:
    include_speaker_notes: bool = False
    pdf: bool = True
    slide_previews: bool = True


ContentBlock: TypeAlias = dict[str, Any]


@dataclass(frozen=True)
class Slide:
    semantic_layout: str
    title: str | None = None
    content_blocks: tuple[ContentBlock, ...] = ()
    images: tuple[Resource, ...] = ()
    speaker_notes: str | None = None
    animation_timeline: Any = None
    id: int | None = None


@dataclass(frozen=True)
class PresentationIr:
    slides: tuple[Slide, ...]
    theme: str | None = None
    master: TemplateRef | None = None
    layouts: tuple[TemplateRef, ...] = ()
    export_policy: ExportPolicy = field(default_factory=ExportPolicy)


@dataclass(frozen=True)
class DeckEnvelope:
    schema_version: str
    kind: str
    document_id: str
    metadata: Metadata
    document: PresentationIr
    template: TemplateRef | None = None
    resources: tuple[Resource, ...] = ()
    validation: tuple[str, ...] = ()
    outputs: tuple[str, ...] = ("pptx", "pdf", "slide-previews")


_BLOCK_TYPES = frozenset({"text", "bullets", "table", "chart", "image"})


def _require(mapping: dict[str, Any], key: str, path: str) -> Any:
    if key not in mapping:
        raise IrValidationError(path, f"missing required key '{key}'")
    return mapping[key]


def _require_str(mapping: dict[str, Any], key: str, path: str) -> str:
    value = _require(mapping, key, path)
    if not isinstance(value, str):
        raise IrValidationError(path, f"'{key}' must be a string, got {type(value).__name__}")
    return value


def parse_content_block(block: Any, path: str) -> ContentBlock:
    if not isinstance(block, dict):
        raise IrValidationError(path, "content block must be an object")
    block_type = _require_str(block, "type", path)
    if block_type not in _BLOCK_TYPES:
        raise IrValidationError(path, f"unknown content block type '{block_type}'")
    if block_type == "bullets":
        items = block.get("items")
        if not isinstance(items, list) or not all(isinstance(i, str) for i in items):
            raise IrValidationError(path, "'items' must be a list of strings")
    if block_type == "table":
        rows = block.get("rows")
        if not isinstance(rows, list) or not all(isinstance(r, list) for r in rows):
            raise IrValidationError(path, "'rows' must be a list of lists")
    return dict(block)


def parse_slide(raw: Any, index: int) -> Slide:
    path = f"document.slides[{index}]"
    if not isinstance(raw, dict):
        raise IrValidationError(path, "slide must be an object")
    layout = _require_str(raw, "semantic_layout", path)
    title = raw.get("title")
    if title is not None and not isinstance(title, str):
        raise IrValidationError(path, "'title' must be a string")
    blocks = tuple(parse_content_block(b, f"{path}.content_blocks[{i}]") for i, b in enumerate(raw.get("content_blocks", [])))
    images = tuple(
        Resource(id=_require_str(r, "id", f"{path}.images[{i}]"), uri=_require_str(r, "uri", f"{path}.images[{i}]"))
        for i, r in enumerate(raw.get("images", []))
    )
    notes = raw.get("speaker_notes")
    if notes is not None and not isinstance(notes, str):
        raise IrValidationError(path, "'speaker_notes' must be a string")
    return Slide(
        semantic_layout=layout,
        title=title,
        content_blocks=blocks,
        images=images,
        speaker_notes=notes,
    )


def parse_presentation(raw: Any) -> PresentationIr:
    path = "document"
    if not isinstance(raw, dict):
        raise IrValidationError(path, "must be an object")
    slides = raw.get("slides")
    if not isinstance(slides, list) or not slides:
        raise IrValidationError(f"{path}.slides", "must be a non-empty list")
    policy_raw = raw.get("export_policy", {})
    policy = ExportPolicy(
        include_speaker_notes=bool(policy_raw.get("include_speaker_notes", False)),
        pdf=bool(policy_raw.get("pdf", True)),
        slide_previews=bool(policy_raw.get("slide_previews", True)),
    )
    return PresentationIr(
        slides=tuple(parse_slide(s, i) for i, s in enumerate(slides)),
        theme=raw.get("theme"),
        export_policy=policy,
    )


def parse_envelope(raw: Any) -> DeckEnvelope:
    if not isinstance(raw, dict):
        raise IrValidationError("$", "envelope must be an object")
    version = _require_str(raw, "schema_version", "$")
    if version != IR_VERSION:
        raise IrValidationError("$.schema_version", f"expected '{IR_VERSION}', got '{version}'")
    kind = _require_str(raw, "kind", "$")
    if kind != DOCUMENT_KIND:
        raise IrValidationError("$.kind", f"expected '{DOCUMENT_KIND}', got '{kind}'")
    metadata_raw = _require(raw, "metadata", "$")
    if not isinstance(metadata_raw, dict):
        raise IrValidationError("$.metadata", "must be an object")
    template_raw = raw.get("template")
    template = None
    if template_raw is not None:
        if not isinstance(template_raw, dict):
            raise IrValidationError("$.template", "must be an object")
        template = TemplateRef(
            uri=_require_str(template_raw, "uri", "$.template"),
            version=str(template_raw.get("version", "0.0.0")),
        )
    return DeckEnvelope(
        schema_version=version,
        kind=kind,
        document_id=_require_str(raw, "document_id", "$"),
        metadata=Metadata(
            title=_require_str(metadata_raw, "title", "$.metadata"),
            author=str(metadata_raw.get("author", "")),
            language=str(metadata_raw.get("language", "zh-CN")),
        ),
        template=template,
        resources=tuple(Resource(id=str(r.get("id", i)), uri=str(r.get("uri", ""))) for i, r in enumerate(raw.get("resources", []))),
        document=parse_presentation(_require(raw, "document", "$")),
        validation=tuple(str(v) for v in raw.get("validation", [])),
        outputs=tuple(str(o) for o in raw.get("outputs", ["pptx", "pdf", "slide-previews"])),
    )


def artifact_stem(document_id: str) -> str:
    """Safe filesystem stem for a document id (document ids may contain
    characters that are invalid in paths, e.g. 'draft:review-deck')."""
    return re.sub(r"[^A-Za-z0-9._-]+", "-", document_id).strip("-") or "deck"


def load_deck_ir(source: str | Path | dict[str, Any]) -> DeckEnvelope:
    """Load and validate a Deck IR document from a JSON file or a mapping."""
    if isinstance(source, dict):
        return parse_envelope(source)
    path = Path(source)
    if not path.is_file():
        raise IrValidationError("$", f"input file not found: {path}")
    try:
        with path.open("r", encoding="utf-8") as handle:
            raw = json.load(handle)
    except json.JSONDecodeError as exc:
        raise IrValidationError("$", f"invalid JSON: {exc}") from exc
    return parse_envelope(raw)
