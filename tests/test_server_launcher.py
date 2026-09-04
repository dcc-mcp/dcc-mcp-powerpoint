from __future__ import annotations

import sys
from pathlib import Path

from dcc_mcp_powerpoint import __version__, server_launcher


def test_serve_command_binds_powerpoint_identity_and_bundled_skills(tmp_path: Path) -> None:
    server = tmp_path / "dcc-mcp-server.exe"
    server.touch()

    command = server_launcher.ServeConfig(server=str(server), mcp_port=43123).command()

    assert command[:4] == [str(server.resolve()), "serve", "--app", "powerpoint"]
    assert command[command.index("--server-name") + 1] == "dcc-mcp-powerpoint"
    assert command[command.index("--app-version") + 1] == __version__
    assert command[command.index("--mcp-port") + 1] == "43123"
    assert Path(command[command.index("--skill-paths") + 1]).is_dir()
    assert "--no-bridge" in command


def test_serve_forwards_the_current_python_runtime(monkeypatch, tmp_path: Path) -> None:
    server = tmp_path / "dcc-mcp-server.exe"
    server.touch()
    observed = {}

    def fake_call(command, *, env):
        observed["command"] = command
        observed["env"] = env
        return 7

    monkeypatch.setattr(server_launcher.subprocess, "call", fake_call)

    result = server_launcher.serve(server_launcher.ServeConfig(server=str(server)))

    assert result == 7
    assert observed["env"][server_launcher.PYTHON_EXECUTABLE_ENV] == sys.executable


def test_missing_server_fails_with_runtime_hint(monkeypatch) -> None:
    monkeypatch.delenv(server_launcher.SERVER_ENV, raising=False)
    monkeypatch.setattr(server_launcher.shutil, "which", lambda _name: None)

    try:
        server_launcher.ServeConfig().command()
    except FileNotFoundError as exc:
        assert "dcc-mcp-server" in str(exc)
        assert server_launcher.SERVER_ENV in str(exc)
    else:
        raise AssertionError("missing server must fail closed")


def test_bundled_tools_run_out_of_process() -> None:
    tool_manifests = sorted(server_launcher.bundled_skills_dir().glob("*/tools.yaml"))

    assert tool_manifests
    for manifest in tool_manifests:
        contents = manifest.read_text(encoding="utf-8")
        assert "affinity: main" not in contents, manifest
        assert "enforce_thread_affinity: true" not in contents, manifest
