"""dcc-mcp-powerpoint — PowerPoint adapter for the DCC-MCP ecosystem.

Application layer over `dcc-mcp-office`: deck generation, slide composition,
review decks from DCC renders, PowerPoint skill packs and the office-host
launcher.

M2 capability surface (authoring + modification):
- deck_ir: the Deck IR contract (mirrors dcc-mcp-office-ir)
- compiler: Deck IR -> PPTX (Open XML implementation, semantic-layer tagged)
- render: PPTX -> PDF + slide previews (desktop COM implementation)
- validate: structural validation reports
- analyze (opt-in module): self-implemented issue analyzer — imported as
  dcc_mcp_powerpoint.analyze, never at package import time (its reader uses
  python-pptx, a dev/test-only dependency)
- layers (opt-in module): smart semantic layers over shapes (native name
  tags ::layer=<name>; visibility/z-order/recolor/assign) — python-pptx
  based, same opt-in policy as compiler/analyze
- edits (opt-in module): ppt-patch/1.0 deck patch engine (text/images/
  slides/layers, all-or-nothing saves) — python-pptx based
- plugins (stdlib-only): plugin registry (discover/validate/run via
  subprocess stdin-JSON contract)
- host_client (opt-in module): stdlib-only JSON-RPC client for the C# host
"""

from .deck_ir import DeckEnvelope, IrValidationError, load_deck_ir
from .render import office_available, render_deck
from .validate import validate_artifacts, validate_envelope

__version__ = "0.2.0"

__all__ = [
    "DeckEnvelope",
    "IrValidationError",
    "__version__",
    "load_deck_ir",
    "office_available",
    "render_deck",
    "validate_artifacts",
    "validate_envelope",
]
