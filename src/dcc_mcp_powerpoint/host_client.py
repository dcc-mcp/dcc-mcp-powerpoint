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
import time
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


def _rpc_work_dir() -> Path:
    """Per-call work directory for file-based stdio redirection.

    Plain mkdir instead of tempfile.mkdtemp: some confined environments
    (agent sandboxes) deny the mkdtemp/chmod paths while allowing ordinary
    directory creation. The caller removes it best-effort.
    """
    base = Path(os.environ.get("DCC_OFFICE_HOST_TMP") or os.environ.get("TEMP") or os.environ.get("TMP") or ".")
    base = base.resolve()
    for _attempt in range(10):
        candidate = base / f"dcc-office-host-{os.getpid()}-{time.time_ns()}"
        try:
            candidate.mkdir()
            return candidate
        except FileExistsError:
            continue
    raise OSError(f"could not create a host work directory under {base}")


def rpc(method: str, params: dict[str, Any], *, app: str = "powerpoint") -> dict[str, Any]:
    """One JSON-RPC exchange with the host over stdin/stdout.

    Stdio is redirected through temporary files instead of OS pipes: the
    wire contract (request JSON on stdin, response JSON on stdout) is
    unchanged, while confined environments that block anonymous pipes can
    still talk to the host, and large host output cannot deadlock.
    """
    binary = find_host_binary()
    if binary is None:
        return {"success": False, "backend": None, "reason": f"OFFICE_HOST_NOT_FOUND: {HOST_EXE} not found"}
    request = {"jsonrpc": "2.0", "id": "req", "method": method, "params": params}
    work = _rpc_work_dir()
    try:
        stdin_path = work / "request.json"
        stdout_path = work / "response.json"
        stderr_path = work / "stderr.txt"
        stdin_path.write_text(json.dumps(request), encoding="utf-8")
        try:
            with stdin_path.open("r", encoding="utf-8") as stdin_file, stdout_path.open("w", encoding="utf-8") as stdout_file, stderr_path.open("w", encoding="utf-8") as stderr_file:
                proc = subprocess.run(
                    [binary, f"--app={app}"],
                    stdin=stdin_file,
                    stdout=stdout_file,
                    stderr=stderr_file,
                    timeout=300,
                    check=False,
                )
        except subprocess.TimeoutExpired as exc:
            return {"success": False, "backend": "office_host", "reason": f"host timed out: {exc}"}
        stdout = stdout_path.read_text(encoding="utf-8")
        stderr = stderr_path.read_text(encoding="utf-8")
    finally:
        shutil.rmtree(work, ignore_errors=True)
    if proc.returncode != 0:
        return {"success": False, "backend": "office_host", "reason": stderr.strip() or "host exited non-zero"}
    try:
        payload = json.loads(stdout)
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
