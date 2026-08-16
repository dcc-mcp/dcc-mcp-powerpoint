"""Runtime purity guard — importing the package must not pull dev deps.

Dependency policy: python-pptx / Pillow / pywin32 are dev/test-only. The
check runs in a clean interpreter so pytest's own imports cannot mask a
violation.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

PROBE = (
    "import sys; "
    "sys.path.insert(0, r'{src}'); "
    "import dcc_mcp_powerpoint; "
    "bad = [m for m in ('pptx', 'PIL', 'win32com') if m in sys.modules]; "
    "assert not bad, 'dev deps imported at package import: ' + repr(bad); "
    "print('pure')"
).format(src=str(ROOT / "src"))


def test_package_import_pulls_no_dev_deps() -> None:
    proc = subprocess.run(
        [sys.executable, "-c", PROBE],
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=120,
        check=False,
    )
    assert proc.returncode == 0, proc.stderr
    assert "pure" in proc.stdout
