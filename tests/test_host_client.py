"""Host client contract tests — stdlib-only JSON-RPC over stdin/stdout."""

from __future__ import annotations

from pathlib import Path

from dcc_mcp_powerpoint._standalone_entry import _is_skill_script_invocation
from dcc_mcp_powerpoint.host_client import find_host_binary, ping, rpc

_NL = chr(10)


def test_find_host_binary_returns_none_without_host(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.delenv("DCC_OFFICE_HOST", raising=False)
    monkeypatch.setenv("PATH", str(tmp_path))
    assert find_host_binary() is None


def test_find_host_binary_env_override(monkeypatch, tmp_path: Path) -> None:
    fake = tmp_path / "host.cmd"
    fake.write_text("@echo off", encoding="utf-8")
    monkeypatch.setenv("DCC_OFFICE_HOST", str(fake))
    assert find_host_binary() == str(fake)


def test_rpc_returns_clean_reason_when_host_missing(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.delenv("DCC_OFFICE_HOST", raising=False)
    monkeypatch.setenv("PATH", str(tmp_path))
    result = ping()
    assert result["success"] is False
    assert "OFFICE_HOST_NOT_FOUND" in result["reason"]


def test_skill_script_invocation_detection(tmp_path: Path) -> None:
    script = tmp_path / "greet.py"
    script.write_text("print('hi')", encoding="utf-8")
    assert _is_skill_script_invocation(["dcc-mcp-powerpoint", str(script)])
    assert not _is_skill_script_invocation(["dcc-mcp-powerpoint", "compile", "--input", "x"])
    assert not _is_skill_script_invocation(["dcc-mcp-powerpoint"])


def test_rpc_round_trip_against_fake_host(monkeypatch, tmp_path: Path) -> None:
    """Envelope contract: request JSON in, response parsed out."""
    fake = tmp_path / "host.cmd"
    response = '{"jsonrpc":"2.0","id":"req","result":{"app":"powerpoint","protocol_version":"office-rpc/1"}}'
    fake.write_text("@echo off" + _NL + "set /p REQ=" + _NL + "echo " + response + _NL, encoding="utf-8")
    monkeypatch.setenv("DCC_OFFICE_HOST", str(fake))
    result = ping()
    assert result["success"] is True
    assert result["backend"] == "office_host"
    assert result["result"]["protocol_version"] == "office-rpc/1"


def test_rpc_surfaces_host_errors(monkeypatch, tmp_path: Path) -> None:
    fake = tmp_path / "host.cmd"
    fake.write_text("@echo off" + _NL + "exit /b 3" + _NL, encoding="utf-8")
    monkeypatch.setenv("DCC_OFFICE_HOST", str(fake))
    result = rpc("office.command.execute", {"capability": "deck.compile"})
    assert result["success"] is False
    assert "host exited non-zero" in result["reason"]
