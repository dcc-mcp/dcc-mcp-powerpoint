"""Host client contract tests — stdlib-only JSON-RPC over stdin/stdout.

Uses tests/_tmphelper.py instead of pytest's tmp_path (DSH sandbox).
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from dcc_mcp_powerpoint._standalone_entry import _is_skill_script_invocation
from dcc_mcp_powerpoint.host_client import (
    batch_convert,
    batch_replace_text,
    find_host_binary,
    handshake,
    ping,
    rpc,
    slide_render,
)

from ._tmphelper import make_tmp_dir, remove_tmp_dir

_NL = chr(10)


@pytest.fixture()
def tmp_path() -> Path:
    directory = make_tmp_dir("host-client")
    yield directory
    remove_tmp_dir(directory)


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


def _fake_run(captured: dict):
    """Fake subprocess.run: record argv, read the request file, write a response."""

    class _Proc:
        returncode = 0

    def _run(argv, stdin, stdout, stderr, timeout, check):
        captured["argv"] = argv
        captured["request"] = json.loads(stdin.read())
        stdout.write(json.dumps({"jsonrpc": "2.0", "id": "req", "result": {"ok": True}}))
        return _Proc()

    return _run


def _patch_host(monkeypatch, tmp_path: Path) -> None:
    fake = tmp_path / "host.cmd"
    fake.write_text("@echo off", encoding="utf-8")
    monkeypatch.setenv("DCC_OFFICE_HOST", str(fake))


def test_rpc_passes_stdio_flag(monkeypatch, tmp_path: Path) -> None:
    _patch_host(monkeypatch, tmp_path)
    captured: dict = {}
    monkeypatch.setattr("dcc_mcp_powerpoint.host_client.subprocess.run", _fake_run(captured))
    result = ping()
    assert result["success"] is True
    assert "--stdio" in captured["argv"]
    assert any(arg.startswith("--app=") for arg in captured["argv"])


def test_handshake_requests_app(monkeypatch, tmp_path: Path) -> None:
    _patch_host(monkeypatch, tmp_path)
    captured: dict = {}
    monkeypatch.setattr("dcc_mcp_powerpoint.host_client.subprocess.run", _fake_run(captured))
    result = handshake("powerpoint")
    assert result["success"] is True
    assert captured["request"]["method"] == "office.host.handshake"
    assert captured["request"]["params"]["requested_app"] == "powerpoint"


def test_slide_render_sends_absolute_paths(monkeypatch, tmp_path: Path) -> None:
    _patch_host(monkeypatch, tmp_path)
    captured: dict = {}
    monkeypatch.setattr("dcc_mcp_powerpoint.host_client.subprocess.run", _fake_run(captured))
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"x")
    result = slide_render(deck, "previews", width=640, height=360)
    assert result["success"] is True
    request = captured["request"]
    assert request["params"]["capability"] == "slide.render"
    payload = request["params"]["input"]
    assert payload["path"] == str(deck.resolve())
    assert payload["output_directory"] == str((Path.cwd() / "previews").resolve())
    assert payload["width"] == 640
    assert payload["height"] == 360


def test_batch_convert_sends_absolute_paths(monkeypatch, tmp_path: Path) -> None:
    _patch_host(monkeypatch, tmp_path)
    captured: dict = {}
    monkeypatch.setattr("dcc_mcp_powerpoint.host_client.subprocess.run", _fake_run(captured))
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"x")
    result = batch_convert([deck], tmp_path / "out")
    assert result["success"] is True
    payload = captured["request"]["params"]["input"]
    assert captured["request"]["params"]["capability"] == "batch.convert"
    assert payload["inputs"] == [str(deck.resolve())]
    assert payload["output_directory"] == str((tmp_path / "out").resolve())
    assert payload["target_format"] == "pdf"


def test_batch_replace_text_defaults_to_dry_run(monkeypatch, tmp_path: Path) -> None:
    _patch_host(monkeypatch, tmp_path)
    captured: dict = {}
    monkeypatch.setattr("dcc_mcp_powerpoint.host_client.subprocess.run", _fake_run(captured))
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"x")
    result = batch_replace_text([deck], [{"find": "a", "replace": "b"}])
    assert result["success"] is True
    request = captured["request"]
    assert request["params"]["capability"] == "batch.replace_text"
    payload = request["params"]["input"]
    assert payload["dry_run"] is True
    assert payload["rules"] == [{"find": "a", "replace": "b"}]
    assert payload["inputs"] == [str(deck.resolve())]
