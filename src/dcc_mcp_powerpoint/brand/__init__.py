"""Brand resolution — packaged DCC-MCP master assets.

Template-first generation (proposal §15.4): decks reference templates by
`brand://` URI; this module resolves the packaged default brand
(`brand://dcc-mcp/*`) to bundled master assets. External brand packages
extend this later without touching the compiler contract.
"""

from __future__ import annotations

from pathlib import Path

PACKAGED_BRAND = "brand://dcc-mcp"
_ASSETS = Path(__file__).resolve().parent / "assets"

# Variant keyed by background luminance: light logo for dark backgrounds.
LOGO_LIGHT = _ASSETS / "dcc-mcp-logo-light.png"
LOGO_DARK = _ASSETS / "dcc-mcp-logo-dark.png"


def resolve_logo(template_uri: str | None, *, dark_background: bool) -> Path | None:
    """Return the packaged master logo for the default brand, else None.

    Contract: unknown brands resolve to None (caller renders without a
    logo) — never to a look-alike asset.
    """
    if not template_uri or not template_uri.startswith(PACKAGED_BRAND):
        return None
    candidate = LOGO_LIGHT if dark_background else LOGO_DARK
    return candidate if candidate.is_file() else None
