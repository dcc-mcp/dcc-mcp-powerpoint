"""Locate, verify and launch the shared C# `office-host` for PowerPoint.

The host binary is distributed by `dcc-mcp-office` through
`dcc-mcp-release-artifacts` (self-contained .NET exe). This module mirrors
the Unity sidecar-launcher precedent: resolve executable, build args, spawn,
then hand the pipe endpoint to the gateway RPC client.

M0: path resolution + argument building only. Spawn/handshake land in M1
together with `dcc-mcp-office-client`.
"""

from __future__ import annotations

import os
import shutil
from dataclasses import dataclass, field
from pathlib import Path

HOST_EXE = "dcc-office-host.exe"
APP_NAME = "powerpoint"
# Named pipe prefix the host listens on: \\.\pipe\dcc-mcp-office-{app}-{sid}-{session}
PIPE_PREFIX = "dcc-mcp-office"


@dataclass
class OfficeHostConfig:
    """Resolution + launch configuration for the PowerPoint office-host."""

    app: str = APP_NAME
    exe: str | None = None
    extra_args: list[str] = field(default_factory=list)

    def resolve_exe(self) -> str | None:
        """Resolve the host executable.

        Order: explicit `exe` → `DCC_OFFICE_HOST` env → PATH → local
        release-artifacts cache.
        """
        if self.exe and Path(self.exe).is_file():
            return self.exe
        env = os.environ.get("DCC_OFFICE_HOST")
        if env and Path(env).is_file():
            return env
        on_path = shutil.which(HOST_EXE)
        if on_path:
            return on_path
        # M1: download/verify from dcc-mcp-release-artifacts, cache under
        # %LOCALAPPDATA%/dcc-mcp/office-host/<version>/.
        return None

    def build_args(self) -> list[str]:
        """Build the host command line: office-host --app=powerpoint ..."""
        exe = self.resolve_exe()
        if exe is None:
            raise FileNotFoundError(
                f"{HOST_EXE} not found; set DCC_OFFICE_HOST or install the "
                "dcc-mcp-office runtime (M1: auto-download from release artifacts)"
            )
        return [exe, f"--app={self.app}", *self.extra_args]

    def pipe_name(self, user_sid: str, session_id: int) -> str:
        """Named pipe for this app in the given user session (proposal §12.1)."""
        return rf"\\.\pipe\{PIPE_PREFIX}-{self.app}-{user_sid}-{session_id}"


def launch(config: OfficeHostConfig | None = None) -> None:
    """M0 placeholder: spawn the host process.

    M1: start process, wait for handshake on the named pipe, register the
    endpoint with the gateway HostRpcClient (namedpipe://powerpoint).
    """
    cfg = config or OfficeHostConfig()
    args = cfg.build_args()
    raise NotImplementedError(
        f"M1: spawning the office-host is not wired yet (args would be: {args})"
    )
