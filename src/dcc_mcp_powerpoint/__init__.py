"""dcc-mcp-powerpoint — PowerPoint adapter for the DCC-MCP ecosystem.

Application layer over `dcc-mcp-office`: deck generation, slide composition,
review decks from DCC renders, PowerPoint skill packs and the office-host
launcher.

M1 capability surface (Open XML backend live, COM backend via render_deck):
- deck_ir: the Deck IR contract (mirrors dcc-mcp-office-ir)
- compiler: Deck IR -> PPTX (Open XML implementation)
- render: PPTX -> PDF + slide previews (desktop COM implementation)
- validate: structural validation reports
- analyze: self-implemented issue analyzer (path addressing, three buckets,
  real font-metric overflow + WCAG contrast)
"""

from .analyze import analyze_deck
from .deck_ir import DeckEnvelope, IrValidationError, load_deck_ir
from .render import office_available, render_deck
from .validate import validate_artifacts, validate_envelope

__version__ = "0.1.0"

__all__ = [
    "DeckEnvelope",
    "IrValidationError",
    "__version__",
    "analyze_deck",
    "load_deck_ir",
    "office_available",
    "render_deck",
    "validate_artifacts",
    "validate_envelope",
]
