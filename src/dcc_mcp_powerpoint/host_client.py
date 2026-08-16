"""dcc-office-host client — stdlib-only JSON-RPC over stdin/stdout.

The C# host (dcc-mcp-office) owns the heavy Office surfaces; this client is
the Python-side contract. Binary resolution order:
1. DCC_OFFICE_HOST env
2. $ORIGIN/lib/dcc-office-host.exe (PyOxidizer standalone layout)
3. dcc-office-host on PATH
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

HOST_EXE = "dcc-office-host.exe"
OFFICE_HOST_ENV = "DCC_OFFICE_HOST"


def find_host_binary() -> str | None:
    """Locate the host executable (see module docstring for the order)."""
    override = os.environ.get(OFFICE_HOST_ENV)
    if override and Path(override).is_file():
        return override
    exe_dir = Path(sys.executable).resolve().parent
    bundled = exe_dir / "lib" / HOST_EXE
    if bundled.is_file():
        return str(bundled)
    on_path = shutil.which(HOST_EXE)
    return on_path


def rpc(method: str, params: dict[str, Any], *, app: str = "powerpoint") -> dict[str, Any]:
    """One JSON-RPC exchange with the host over stdin/stdout."""
    binary = find_host_binary()
    if binary is None:
        return {"success": False, "backend": None, "reason": f"OFFICE_HOST_NOT_FOUND: {HOST_EXE} not found"}
    request = {"jsonrpc": "2.0", "id": "req", "method": method, "params": params}
    proc = subprocess.run(
        [binary, f"--app={app}"],
        input=json.dumps(request),
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=300,
        check=False,
    )
    if proc.returncode != 0:
        return {"success": False, "backend": "office_host", "reason": proc.stderr.strip() or "host exited non-zero"}
    try:
        payload = json.loads(proc.stdout)
    except json.JSONDecodeError as exc:
        return {"success": False, "backend": "office_host", "reason": f"host output not JSON: {exc}"}
    result = payload.get("result")
    if result is None:
        return {"success": False, "backend": "office_host", "reason": str(payload.get("error", "empty result"))}
    return {"success": True, "backend": "office_host", "result": result}


def ping() -> dict[str, Any]:
    return rpc("office.host.ping", {})


def compile_deck(ir_path: str | Path, output_path: str | Path) -> dict[str, Any]:
    return rpc("office.command.execute", {"capability": "deck.compile", "input": {"ir": str(ir_path), "output": str(output_path)}})


def inspect_deck(pptx_path: str | Path) -> dict[str, Any]:
    return rpc("office.command.execute", {"capability": "document.inspect", "input": {"path": str(pptx_path)}})
