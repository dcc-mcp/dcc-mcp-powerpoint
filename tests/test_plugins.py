"""Plugin registry tests — discovery/validation with real files, execution via monkeypatch.

Uses tests/_tmphelper.py instead of pytest's tmp_path (DSH sandbox).
"""

from __future__ import annotations

import json
import os
from pathlib import Path

import pytest

from dcc_mcp_powerpoint.plugins import (
    DEFAULT_PLUGIN_ROOT,
    MAX_TIMEOUT_MS,
    PluginValidationError,
    discover_plugins,
    load_manifest,
    plugin_roots,
    resolve_plugin,
    run_plugin,
)

from ._tmphelper import make_tmp_dir, remove_tmp_dir


def _make_plugin(root: Path, name: str = "my-plugin", **overrides) -> Path:
    directory = root / name
    directory.mkdir(parents=True, exist_ok=True)
    manifest = {
        "name": name,
        "version": "1.0.0",
        "description": f"{name} demo plugin",
        "script": "run.py",
    }
    manifest.update(overrides)
    (directory / "plugin.json").write_text(json.dumps(manifest), encoding="utf-8")
    (directory / "run.py").write_text('print("{}")\n', encoding="utf-8")
    return directory


@pytest.fixture()
def tmp_root() -> Path:
    directory = make_tmp_dir("plugins")
    yield directory
    remove_tmp_dir(directory)


def test_plugin_roots_defaults(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("DCC_POWERPOINT_PLUGIN_PATH", raising=False)
    roots = plugin_roots()
    assert DEFAULT_PLUGIN_ROOT in roots


def test_plugin_roots_env_override(monkeypatch: pytest.MonkeyPatch, tmp_root: Path) -> None:
    monkeypatch.setenv("DCC_POWERPOINT_PLUGIN_PATH", str(tmp_root))
    roots = plugin_roots()
    assert Path(roots[0]) == tmp_root


def test_load_manifest_valid(tmp_root: Path) -> None:
    directory = _make_plugin(tmp_root, "valid-one")
    manifest = load_manifest(directory)
    assert manifest["name"] == "valid-one"
    assert manifest["script"] == "run.py"


@pytest.mark.parametrize(
    ("overrides", "match"),
    [
        ({"name": "Bad Name!"}, "slug"),
        ({"name": "no-description"}, None),
        ({"script": "missing.py"}, "script not found"),
        ({"script": "../escape.py"}, "escapes the plugin directory"),
        ({"timeout_ms": 0}, "timeout_ms"),
        ({"timeout_ms": MAX_TIMEOUT_MS + 1}, "timeout_ms"),
    ],
)
def test_load_manifest_invalid(tmp_root: Path, overrides: dict, match: str | None) -> None:
    directory = _make_plugin(tmp_root, **overrides)
    if overrides.get("name") == "no-description":
        raw = json.loads((directory / "plugin.json").read_text(encoding="utf-8"))
        raw.pop("description")
        (directory / "plugin.json").write_text(json.dumps(raw), encoding="utf-8")
    with pytest.raises(PluginValidationError, match=match or "description"):
        load_manifest(directory)


def test_load_manifest_missing_file(tmp_root: Path) -> None:
    directory = tmp_root / "empty"
    directory.mkdir()
    with pytest.raises(PluginValidationError, match="plugin.json not found"):
        load_manifest(directory)


def test_discover_plugins_and_duplicates(tmp_root: Path) -> None:
    root = tmp_root / "plugins"
    _make_plugin(root, "alpha")
    _make_plugin(root / "other", "alpha")  # duplicate name in a second root
    bad = root / "broken"
    bad.mkdir()
    (bad / "plugin.json").write_text("{not json", encoding="utf-8")
    report = discover_plugins(f"{root}{os.pathsep}{root / 'other'}")
    names = [p["name"] for p in report["plugins"]]
    assert names == ["alpha"]
    assert any("duplicate plugin name" in e["message"] for e in report["errors"])
    assert any("invalid JSON" in e["message"] for e in report["errors"])


def test_resolve_plugin_by_name_and_dir(tmp_root: Path) -> None:
    root = tmp_root / "plugins"
    directory = _make_plugin(root, "alpha")
    by_name = resolve_plugin("alpha", paths=str(root))
    by_dir = resolve_plugin(str(directory))
    assert by_name["dir"] == by_dir["dir"] == str(directory.resolve())


def test_resolve_plugin_unknown(tmp_root: Path) -> None:
    with pytest.raises(PluginValidationError, match="not found"):
        resolve_plugin("nope", paths=str(tmp_root))


class _FakeProc:
    def __init__(self, returncode: int = 0):
        self.returncode = returncode


def _fake_run(captured: dict):
    def _run(argv, stdin, stdout, stderr, timeout, check):
        captured["argv"] = argv
        captured["payload"] = json.loads(stdin.read())
        captured["timeout"] = timeout
        stdout.write(json.dumps({"success": True, "answer": 42}))
        return _FakeProc(0)

    return _run


def test_run_plugin_contract(monkeypatch: pytest.MonkeyPatch, tmp_root: Path) -> None:
    _make_plugin(tmp_root, "contract")
    captured: dict = {}
    monkeypatch.setattr("dcc_mcp_powerpoint.plugins.subprocess.run", _fake_run(captured))
    result = run_plugin("contract", {"n": 3}, {"pptx": "deck.pptx"}, paths=str(tmp_root))
    assert result["success"] is True
    assert result["answer"] == 42
    assert captured["argv"][-1].endswith(str(Path("run.py")))
    assert captured["payload"] == {"context": {"pptx": "deck.pptx"}, "params": {"n": 3}}
    assert captured["timeout"] == 120.0


def test_run_plugin_timeout_cap(monkeypatch: pytest.MonkeyPatch, tmp_root: Path) -> None:
    directory = _make_plugin(tmp_root, "slow", timeout_ms=30000)
    captured: dict = {}
    monkeypatch.setattr("dcc_mcp_powerpoint.plugins.subprocess.run", _fake_run(captured))
    run_plugin(str(directory), None, None)
    assert captured["timeout"] == 30.0


def test_run_plugin_nonzero_exit(monkeypatch: pytest.MonkeyPatch, tmp_root: Path) -> None:
    directory = _make_plugin(tmp_root, "crash")

    def _run(argv, stdin, stdout, stderr, timeout, check):
        stdin.read()
        stderr.write("boom")
        return _FakeProc(1)

    monkeypatch.setattr("dcc_mcp_powerpoint.plugins.subprocess.run", _run)
    result = run_plugin(str(directory), None, None)
    assert result["success"] is False
    assert "exited with code 1" in result["reason"]
    assert "boom" in result["stderr"]


def test_run_plugin_non_json_stdout(monkeypatch: pytest.MonkeyPatch, tmp_root: Path) -> None:
    directory = _make_plugin(tmp_root, "talkative")

    def _run(argv, stdin, stdout, stderr, timeout, check):
        stdin.read()
        stdout.write("not json at all")
        return _FakeProc(0)

    monkeypatch.setattr("dcc_mcp_powerpoint.plugins.subprocess.run", _run)
    result = run_plugin(str(directory), None, None)
    assert result["success"] is False
    assert "not JSON" in result["reason"]
