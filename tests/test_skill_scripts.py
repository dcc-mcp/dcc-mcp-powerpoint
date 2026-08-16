"""Skill script integration tests — run the scripts like the gateway would.

The gateway executes skills/scripts/<name>.py as a subprocess feeding stdin
JSON (dcc-mcp-core execute_script convention); these tests pin that contract.
COM stays at the outermost boundary: the test process itself never touches
PowerPoint, and a COM regression (e.g. the 0x80010108 disconnect noise)
fails the test via the stderr assertion.
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
DECK_SCRIPT = ROOT / "src/dcc_mcp_powerpoint/skills/powerpoint-deck/scripts/generate_deck.py"
VALIDATE_SCRIPT = ROOT / "src/dcc_mcp_powerpoint/skills/powerpoint-deck/scripts/validate_deck.py"
REVIEW_SCRIPT = ROOT / "src/dcc_mcp_powerpoint/skills/powerpoint-review/scripts/review_deck_from_renders.py"
EXAMPLE = ROOT / "examples/dcc_mcp_framework_intro.json"
MANIFEST = ROOT / "examples/shots_manifest.example.json"


def _run_script(script: Path, params: dict) -> tuple[dict, str]:
    proc = subprocess.run(
        [sys.executable, str(script)],
        input=json.dumps(params),
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=300,
        check=False,
    )
    assert proc.returncode == 0, proc.stderr
    assert "fatal" not in proc.stderr.lower(), proc.stderr
    return json.loads(proc.stdout), proc.stderr


def test_generate_deck_script_without_render(tmp_path: Path) -> None:
    result, _ = _run_script(
        DECK_SCRIPT,
        {"input": str(EXAMPLE), "output_dir": str(tmp_path), "render": False},
    )
    assert result["success"], result
    pptx = tmp_path / "draft-dcc-mcp-framework-intro.pptx"
    assert pptx.is_file()
    assert result["context"]["backend"] == "openxml"
    assert result["context"]["validation"]["ok"]


def test_validate_deck_script(tmp_path: Path) -> None:
    result, _ = _run_script(VALIDATE_SCRIPT, {"input": str(EXAMPLE)})
    assert result["success"], result
    assert result["context"]["ok"]


def test_validate_artifacts_script(tmp_path: Path) -> None:
    _run_script(DECK_SCRIPT, {"input": str(EXAMPLE), "output_dir": str(tmp_path), "render": False})
    result, _ = _run_script(VALIDATE_SCRIPT, {"input": str(tmp_path)})
    assert result["success"], result


def test_generate_deck_script_with_com_render(tmp_path: Path) -> None:
    result, _ = _run_script(
        DECK_SCRIPT,
        {"input": str(EXAMPLE), "output_dir": str(tmp_path), "render": True, "previews": True},
    )
    if result["context"]["backend"] != "desktop_com":
        pytest.skip("PowerPoint COM not available on this machine")
    assert result["success"], result
    pdf = tmp_path / "draft-dcc-mcp-framework-intro.pdf"
    assert pdf.is_file()
    previews = list((tmp_path / "previews").glob("slide-*.png"))
    assert len(previews) == 14
    for png in previews:
        assert _png_size(png) == (1920, 1080), f"unexpected preview size: {png.name}"
        coverage, bleed = _pixel_qa(png)
        assert 0.005 <= coverage <= 0.85, f"{png.name}: implausible content coverage {coverage:.3f}"
        assert bleed < 0.005, f"{png.name}: content bleeds into the slide edge ({bleed:.4f})"


def _pixel_qa(path: Path) -> tuple[float, float]:
    """Content coverage + edge-bleed fraction, background-adaptive.

    The 12px border band defines the slide background (its histogram mode):
    layouts keep the edges clear, so the border is pure background. Any
    pixel deviating by more than 12 luminance levels counts as content.
    coverage = content fraction of the full frame vs that background;
    bleed = content fraction inside the border band itself. Works for dark
    and light decks alike and stays correct when content covers ~half the
    slide.
    """
    from PIL import Image

    def _content_frac(hist: list[int], total: int, bg: int) -> float:
        content = sum(count for level, count in enumerate(hist) if abs(level - bg) > 12)
        return content / total

    with Image.open(path) as img:
        gray = img.convert("L")
        width, height = gray.size
        total = width * height
        bands = [
            (0, 0, width, 12),
            (0, height - 12, width, height),
            (0, 0, 12, height),
            (width - 12, 0, width, height),
        ]
        # Accumulate element-wise (histograms are 256-length lists; += on a
        # list extends it instead of adding element-wise).
        band_hist = [0] * 256
        band_pixels = 0
        for box in bands:
            crop = gray.crop(box)
            band_pixels += crop.size[0] * crop.size[1]
            for level, count in enumerate(crop.histogram()):
                band_hist[level] += count
        bg = max(range(256), key=band_hist.__getitem__)
        coverage = _content_frac(gray.histogram(), total, bg)
        bleed = _content_frac(band_hist, band_pixels, bg)
    return coverage, bleed


def _png_size(path: Path) -> tuple[int, int]:
    """Read IHDR width/height from the PNG header (no extra dependency)."""
    header = path.read_bytes()[:24]
    assert header[:8] == b"\x89PNG\r\n\x1a\n", "not a PNG"
    width = int.from_bytes(header[16:20], "big")
    height = int.from_bytes(header[20:24], "big")
    return width, height


def test_review_deck_script(tmp_path: Path) -> None:
    result, _ = _run_script(
        REVIEW_SCRIPT,
        {"input": str(MANIFEST), "output_dir": str(tmp_path), "render": False},
    )
    assert result["success"], result
    deck = tmp_path / "draft-review-2026-08-16.pptx"
    assert deck.is_file()
    assert result["context"]["artifacts_ok"]["ok"]
