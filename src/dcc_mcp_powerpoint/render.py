"""COM renderer — PPTX → PDF + per-slide PNG previews via PowerPoint.

The desktop COM backend (proposal §6.2): native finalization, export and
rendering. Contract-first: this module never claims success silently —
when Office is unavailable it reports backend=None with an explicit reason
(OFFICE_APP_NOT_INSTALLED), never a fake artifact.
"""

from __future__ import annotations

import functools
import logging
import time
from pathlib import Path
from typing import Any

POWERPNT_EXE_CANDIDATES = (
    Path("C:/Program Files/Microsoft Office/root/Office16/POWERPNT.EXE"),
    Path("C:/Program Files (x86)/Microsoft Office/root/Office16/POWERPNT.EXE"),
)

# PowerPoint file-format constant for PDF (ppSaveAsPDF).
_PP_SAVE_AS_PDF = 32

logger = logging.getLogger(__name__)


@functools.lru_cache(maxsize=1)
def office_available() -> bool:
    """True when PowerPoint is installed and COM dispatch works (cached)."""
    if not any(p.is_file() for p in POWERPNT_EXE_CANDIDATES):
        return False
    try:
        import win32com.client

        app = win32com.client.DispatchEx("PowerPoint.Application")
        try:
            return bool(getattr(app, "Version", ""))
        finally:
            # Dedicated probe instance: quit immediately, never leak a
            # process and never touch a user-started PowerPoint (§8.3).
            app.Quit()
    except Exception as exc:  # noqa: BLE001 — any COM failure means "unavailable"
        logger.debug("PowerPoint COM probe failed: %s", exc)
        return False


def render_deck(pptx_path: str | Path, out_dir: str | Path, *, pdf: bool = True, previews: bool = True) -> dict[str, Any]:
    """Render a compiled deck with the desktop COM backend.

    Returns a result context: backend, pdf path, preview paths, office
    version — or an explicit unavailable reason. Never degrades silently.
    """
    # COM runs in PowerPoint's process: relative paths would resolve against
    # PowerPoint's working directory (usually System32), not ours — always
    # hand absolute paths to the COM surface (0x80070003 lesson).
    pptx = Path(pptx_path).resolve()
    if not pptx.is_file():
        return {"success": False, "backend": None, "reason": f"input not found: {pptx}"}
    out = Path(out_dir).resolve()
    out.mkdir(parents=True, exist_ok=True)

    if not office_available():
        # Self-implemented HTML-engine preview backend is roadmap (see
        # learnings.md — OfficeCLI research, 2026-08-16). Until then, no
        # Office means an explicit unavailable reason, never a fake artifact.
        return {
            "success": False,
            "backend": None,
            "reason": "OFFICE_APP_NOT_INSTALLED: PowerPoint COM is unavailable; PPTX compiled, PDF/previews skipped",
        }

    import win32com.client

    context: dict[str, Any] = {"backend": "desktop_com"}
    app = None
    pres = None
    try:
        # DispatchEx: dedicated instance. Never attach to (or quit) a
        # PowerPoint the user started themselves (proposal §8.3).
        app = win32com.client.DispatchEx("PowerPoint.Application")
        version = getattr(app, "Version", "unknown")
        context["office_version"] = version
        pres = app.Presentations.Open(str(pptx), ReadOnly=True, WithWindow=False)
        if pdf:
            pdf_path = out / f"{pptx.stem}.pdf"
            pres.SaveAs(str(pdf_path), _PP_SAVE_AS_PDF)
            context["pdf"] = str(pdf_path)
        if previews:
            preview_dir = out / "previews"
            preview_dir.mkdir(exist_ok=True)
            paths = []
            for index, slide in enumerate(pres.Slides, start=1):
                png = preview_dir / f"slide-{index:02d}.png"
                slide.Export(str(png), "PNG", 1920, 1080)
                paths.append(str(png))
            context["previews"] = paths
        context["success"] = True
    except Exception as exc:  # noqa: BLE001 — report, never swallow
        context = {"success": False, "backend": "desktop_com", "reason": f"OFFICE_RENDER_TIMEOUT/RENDER_FAILED: {exc}"}
    finally:
        # Best-effort cleanup — never mask the real result.
        try:
            if pres is not None:
                pres.Close()
        except Exception as exc:  # noqa: BLE001
            logger.debug("pres.Close failed: %s", exc)
        try:
            if app is not None:
                app.Quit()
        except Exception as exc:  # noqa: BLE001
            logger.debug("app.Quit failed: %s", exc)
        time.sleep(0.3)
    return context
