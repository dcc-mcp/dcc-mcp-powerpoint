"""dcc-office-host client — stdlib-only JSON-RPC over stdin/stdout.

The C# host (dcc-mcp-office) owns the heavy Office surfaces; this client is
the Python-side contract. Binary resolution order:
1. DCC_OFFICE_HOST env
2. $ORIGIN/lib/dcc-office-host.exe (PyOxidizer standalone layout)
3. dcc-office-host on PATH

Capability surface (host >= v0.2.0): ping, handshake (capability manifest),
deck.compile, document.inspect, slide.render, batch.convert,
batch.replace_text. All client paths are absolute — the host's COM backends
resolve relative paths against the Office process working directory.
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


def _abs(path: str | Path) -> str:
    """Absolute path for every path handed to the host.

    The host's COM backends resolve relative paths against the Office
    process working directory (usually System32), so relative input is
    guaranteed to fail with 0x80070003 — normalize on the client side and
    never pass relative paths (belt and suspenders with the host fix).
    """
    return str(Path(path).resolve())


def rpc(method: str, params: dict[str, Any], *, app: str = "powerpoint") -> dict[str, Any]:
    """One JSON-RPC exchange with the host over stdin/stdout.

    Stdio is redirected through temporary files instead of OS pipes: the
    wire contract (request JSON on stdin, response JSON on stdout) is
    unchanged, while confined environments that block anonymous pipes can
    still talk to the host, and large host output cannot deadlock.

    The host is invoked with --stdio: since v0.2.0 the host defaults to
    its named-pipe server and requires the flag for the stdin/stdout loop;
    older hosts ignore the unknown flag and stay on stdio.
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
                    [binary, f"--app={app}", "--stdio"],
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


def handshake(app: str = "powerpoint") -> dict[str, Any]:
    """office.host.handshake — protocol + capability manifest (v0.2.0+)."""
    return rpc("office.host.handshake", {"requested_app": app}, app=app)


def compile_deck(ir_path: str | Path, output_path: str | Path) -> dict[str, Any]:
    return rpc(
        "office.command.execute",
        {"capability": "deck.compile", "input": {"ir": _abs(ir_path), "output": _abs(output_path)}},
    )


def inspect_deck(pptx_path: str | Path, *, backend: str | None = None) -> dict[str, Any]:
    input_payload = {"path": _abs(pptx_path)}
    if backend:
        input_payload["backend"] = backend
    return rpc("office.command.execute", {"capability": "document.inspect", "input": input_payload})


def slide_render(pptx_path: str | Path, output_dir: str | Path, *, width: int = 1280, height: int = 720) -> dict[str, Any]:
    """slide.render — PNG previews per slide + shape overflow detection (v0.2.0+)."""
    return rpc(
        "office.command.execute",
        {
            "capability": "slide.render",
            "input": {
                "path": _abs(pptx_path),
                "output_directory": _abs(output_dir),
                "width": width,
                "height": height,
            },
        },
    )


def batch_convert(inputs: list[str | Path], output_dir: str | Path, *, target_format: str = "pdf") -> dict[str, Any]:
    """batch.convert — high-fidelity PDF per file via the COM backend (v0.2.0+)."""
    return rpc(
        "office.command.execute",
        {
            "capability": "batch.convert",
            "input": {
                "inputs": [_abs(p) for p in inputs],
                "output_directory": _abs(output_dir),
                "target_format": target_format,
            },
        },
    )


def batch_replace_text(
    inputs: list[str | Path],
    rules: list[dict[str, str]],
    *,
    scope: list[str] | None = None,
    dry_run: bool = True,
) -> dict[str, Any]:
    """batch.replace_text — dry-run/commit text replacement via COM (v0.2.0+)."""
    input_payload: dict[str, Any] = {
        "inputs": [_abs(p) for p in inputs],
        "rules": rules,
        "dry_run": dry_run,
    }
    if scope is not None:
        input_payload["scope"] = scope
    return rpc("office.command.execute", {"capability": "batch.replace_text", "input": input_payload})
