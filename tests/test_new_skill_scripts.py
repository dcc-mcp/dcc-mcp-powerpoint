"""Skill script integration tests for the new skills (layers/edit/plugins).

Same convention as test_skill_scripts.py: run the scripts like the gateway
would (subprocess + stdin JSON). COM stays at the outermost boundary; these
tests never touch PowerPoint.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "src/dcc_mcp_powerpoint/skills"
EXAMPLES = ROOT / "examples"

ENV = dict(os.environ)
SRC = str(ROOT / "src")
ENV["PYTHONPATH"] = SRC + (os.pathsep + ENV["PYTHONPATH"] if ENV.get("PYTHONPATH") else "")

# Sandbox-safe temp dirs (DSH sandbox denies pytest's extended-path tmp_path).
from ._tmphelper import make_tmp_dir


def _run_script(script: Path, params: dict | None = None, *, stdin: str | None = None) -> tuple[dict, str]:
    proc = subprocess.run(
        [sys.executable, str(script)],
        input=stdin if stdin is not None else json.dumps(params or {}),
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=300,
        check=False,
        env=ENV,
    )
    assert proc.returncode == 0, proc.stderr
    return json.loads(proc.stdout), proc.stderr


def _make_deck(directory: Path) -> Path:
    from dcc_mcp_powerpoint.compiler import compile_deck
    from dcc_mcp_powerpoint.deck_ir import load_deck_ir

    return compile_deck(load_deck_ir(EXAMPLES / "dcc_mcp_framework_intro.json"), directory / "deck.pptx")


@pytest.fixture()
def tmp_dir() -> Path:
    return make_tmp_dir("skill-scripts")


LAYER_DIR = SCRIPTS / "powerpoint-layers/scripts"
EDIT_DIR = SCRIPTS / "powerpoint-edit/scripts"
PLUGIN_DIR = SCRIPTS / "powerpoint-plugins/scripts"


def test_layer_list_script(tmp_dir: Path) -> None:
    deck = _make_deck(tmp_dir)
    result, _ = _run_script(LAYER_DIR / "layer_list.py", {"input": str(deck)})
    assert result["success"]
    assert result["context"]["layers"]
    assert result["context"]["untagged"] == []


def test_layer_set_visibility_script(tmp_dir: Path) -> None:
    deck = _make_deck(tmp_dir)
    out = tmp_dir / "hidden.pptx"
    result, _ = _run_script(LAYER_DIR / "layer_set_visibility.py", {"input": str(deck), "layer": "decoration", "visible": False, "output": str(out)})
    assert result["success"]
    assert out.is_file()
    assert result["context"]["visibility"]["changed"] > 0


def test_layer_reorder_script(tmp_dir: Path) -> None:
    deck = _make_deck(tmp_dir)
    out = tmp_dir / "reordered.pptx"
    result, _ = _run_script(LAYER_DIR / "layer_reorder.py", {"input": str(deck), "layer": "content", "position": "front", "output": str(out)})
    assert result["success"]
    assert result["context"]["reorder"]["slides_moved"] > 0


def test_layer_recolor_script(tmp_dir: Path) -> None:
    deck = _make_deck(tmp_dir)
    out = tmp_dir / "recolored.pptx"
    result, _ = _run_script(LAYER_DIR / "layer_recolor.py", {"input": str(deck), "layer": "accent", "fill": "E4572E", "output": str(out)})
    assert result["success"]
    assert result["context"]["recolor"]["details"]["fill"] > 0


def test_layer_assign_script(tmp_dir: Path) -> None:
    deck = _make_deck(tmp_dir)
    result, _ = _run_script(LAYER_DIR / "layer_assign.py", {"input": str(deck), "layer": "media", "select": {"shape": "Picture"}})
    assert result["success"]
    assert result["context"]["assign"]["tagged"] > 0


def test_slide_tree_script(tmp_dir: Path) -> None:
    deck = _make_deck(tmp_dir)
    result, _ = _run_script(EDIT_DIR / "slide_tree.py", {"input": str(deck), "slide": 2})
    assert result["success"]
    assert result["context"]["slides"][0]["slide"] == 2


def test_edit_deck_script(tmp_dir: Path) -> None:
    deck = _make_deck(tmp_dir)
    patch_path = tmp_dir / "patch.json"
    patch = {
        "schema_version": "ppt-patch/1.0",
        "operations": [
            {"op": "set_text", "select": {"slide": 2, "layer": "header"}, "text": "EDITED"},
            {"op": "add_slide", "slide_ir": {"semantic_layout": "bullets", "title": "New", "content_blocks": [{"type": "bullets", "items": ["a"]}]}},
        ],
    }
    patch_path.write_text(json.dumps(patch), encoding="utf-8")
    out = tmp_dir / "v2.pptx"
    result, _ = _run_script(EDIT_DIR / "edit_deck.py", {"input": str(deck), "patch": str(patch_path), "output": str(out)})
    assert result["success"]
    assert out.is_file()
    assert result["context"]["slides_after"] == 17


def test_edit_deck_script_reports_failure_without_saving(tmp_dir: Path) -> None:
    deck = _make_deck(tmp_dir)
    original = deck.read_bytes()
    patch_path = tmp_dir / "bad.json"
    patch_path.write_text(json.dumps({"schema_version": "ppt-patch/1.0", "operations": [{"op": "delete_slide", "slide": 99}]}), encoding="utf-8")
    proc = subprocess.run(
        [sys.executable, str(EDIT_DIR / "edit_deck.py")],
        input=json.dumps({"input": str(deck), "patch": str(patch_path)}),
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=300,
        check=False,
        env=ENV,
    )
    assert proc.returncode == 1
    result = json.loads(proc.stdout)
    assert result["success"] is False
    assert "operation 0 failed" in result["message"]
    assert deck.read_bytes() == original


def test_plugin_list_script(tmp_dir: Path) -> None:
    root = tmp_dir / "plugins"
    plugin_dir = root / "demo"
    plugin_dir.mkdir(parents=True)
    (plugin_dir / "plugin.json").write_text(
        json.dumps({"name": "demo", "version": "1.0.0", "description": "demo", "script": "run.py"}), encoding="utf-8"
    )
    (plugin_dir / "run.py").write_text('print(json.dumps({"success": True}))\\n', encoding="utf-8")
    result, _ = _run_script(PLUGIN_DIR / "plugin_list.py", {"paths": str(root)})
    assert result["success"]
    assert [p["name"] for p in result["context"]["plugins"]] == ["demo"]


def test_plugin_run_script(tmp_dir: Path) -> None:
    root = tmp_dir / "plugins"
    plugin_dir = root / "echo"
    plugin_dir.mkdir(parents=True)
    (plugin_dir / "plugin.json").write_text(
        json.dumps({"name": "echo", "version": "1.0.0", "description": "echo params", "script": "run.py"}), encoding="utf-8"
    )
    (plugin_dir / "run.py").write_text(
        "import json, sys\npayload = json.load(sys.stdin)\nprint(json.dumps({'success': True, 'params': payload['params'], 'pptx': payload['context'].get('pptx')}))\n",
        encoding="utf-8",
    )
    result, _ = _run_script(PLUGIN_DIR / "plugin_run.py", {"name": "echo", "paths": str(root), "pptx": "deck.pptx", "params": {"n": 7}})
    assert result["success"]
    assert result["context"]["params"] == {"n": 7}
    assert result["context"]["pptx"] == "deck.pptx"
