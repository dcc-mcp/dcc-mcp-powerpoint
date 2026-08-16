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
    assert len(previews) == 12
    for png in previews:
        assert _png_size(png) == (1920, 1080), f"unexpected preview size: {png.name}"


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
