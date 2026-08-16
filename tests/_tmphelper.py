"""Local temp dir helper — sandbox-safe alternative to pytest's tmp_path.

DSH sandbox note: pytest's tmp_path machinery resolves to Windows extended
(\\?\\) paths and those are denied by the file sandbox. Tests in this
repo that can't rely on tmp_path use this helper instead: directories are
created under .test-tmp with normal paths. CI is unaffected either way.
"""

from __future__ import annotations

import os
import shutil
import time
from pathlib import Path

TEST_TMP = Path(__file__).resolve().parent / ".test-tmp"


def make_tmp_dir(name: str) -> Path:
    directory = TEST_TMP / f"{name}-{os.getpid()}-{time.time_ns()}"
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def remove_tmp_dir(directory: Path) -> None:
    # Best effort: CI cleans up; the sandbox may keep the directory.
    shutil.rmtree(directory, ignore_errors=True)
