"""Plugin registry — discover, validate and run user-installed PPT plugins.

A plugin is a directory containing a plugin.json manifest plus a script:

    plugins/
      my-splitter/
        plugin.json      # name/version/description/script/interpreter/input_schema
        splitter.py      # reads stdin JSON, prints one JSON result to stdout

The manifest contract:

    {
      "name": "my-splitter",          # slug, unique across discovered dirs
      "version": "1.0.0",             # free-form string
      "description": "...",           # one line for agent tool descriptions
      "script": "splitter.py",        # relative to the manifest directory
      "interpreter": "python",        # optional; default: the running Python
      "input_schema": { ... },        # optional JSON Schema for params
      "timeout_ms": 120000            # optional, capped at 600000
    }

Discovery roots (explicit only — never auto-downloaded, never scanned
outside them):

    1. DCC_POWERPOINT_PLUGIN_PATH (os.pathsep-separated directories)
    2. ~/.dcc-mcp/powerpoint/plugins

Execution contract: the script is launched as a subprocess with stdin JSON
{"context": {...}, "params": {...}} and must print one JSON object to stdout
containing at least {"success": bool}. stderr is captured for diagnostics.
Plugins are user-installed code and run with the user's permissions; the
registry never fetches, installs, or imports plugin code itself.

This module is stdlib-only and safe to import at package import time.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path
from typing import Any

DEFAULT_PLUGIN_ROOT = Path.home() / ".dcc-mcp" / "powerpoint" / "plugins"
PLUGIN_PATH_ENV = "DCC_POWERPOINT_PLUGIN_PATH"
MAX_TIMEOUT_MS = 600_000

_REQUIRED_MANIFEST_KEYS = ("name", "version", "description", "script")
_NAME_RE_CHARS = set("abcdefghijklmnopqrstuvwxyz0123456789-_")


class PluginValidationError(ValueError):
    """A plugin directory or manifest violated the plugin contract."""


def plugin_roots(paths: str | None = None) -> list[Path]:
    """Discovery roots: explicit dirs from env/argument + the user default."""
    roots: list[Path] = []
    raw = paths if paths is not None else os.environ.get(PLUGIN_PATH_ENV, "")
    for part in raw.split(os.pathsep):
        if part.strip():
            roots.append(Path(part.strip()))
    roots.append(DEFAULT_PLUGIN_ROOT)
    return roots


def _validate_name(name: Any) -> str:
    if not isinstance(name, str) or not name:
        raise PluginValidationError("manifest 'name' must be a non-empty string")
    if any(char not in _NAME_RE_CHARS for char in name.lower()):
        raise PluginValidationError(f"manifest name {name!r} must be a slug [a-z0-9-_]")
    return name


def load_manifest(plugin_dir: str | Path) -> dict[str, Any]:
    """Load and validate one plugin.json manifest; returns the raw manifest."""
    directory = Path(plugin_dir)
    manifest_path = directory / "plugin.json"
    if not manifest_path.is_file():
        raise PluginValidationError(f"{directory}: plugin.json not found")
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise PluginValidationError(f"{manifest_path}: invalid JSON: {exc}") from exc
    if not isinstance(manifest, dict):
        raise PluginValidationError(f"{manifest_path}: manifest must be an object")
    missing = [key for key in _REQUIRED_MANIFEST_KEYS if key not in manifest]
    if missing:
        raise PluginValidationError(f"{manifest_path}: missing keys: {missing}")
    _validate_name(manifest.get("name"))
    if not isinstance(manifest.get("description"), str) or not manifest["description"].strip():
        raise PluginValidationError(f"{manifest_path}: 'description' must be a non-empty string")
    script = directory / str(manifest["script"])
    resolved = script.resolve()
    if not resolved.is_relative_to(directory.resolve()):
        raise PluginValidationError(f"{manifest_path}: script escapes the plugin directory")
    if not resolved.is_file():
        raise PluginValidationError(f"{manifest_path}: script not found: {manifest['script']}")
    schema = manifest.get("input_schema")
    if schema is not None and not isinstance(schema, dict):
        raise PluginValidationError(f"{manifest_path}: 'input_schema' must be an object")
    timeout = manifest.get("timeout_ms")
    if timeout is not None and (not isinstance(timeout, int) or not 0 < timeout <= MAX_TIMEOUT_MS):
        raise PluginValidationError(f"{manifest_path}: 'timeout_ms' must be an int in 1..{MAX_TIMEOUT_MS}")
    return manifest


def discover_plugins(paths: str | None = None) -> dict[str, Any]:
    """Scan the discovery roots; returns {plugins, errors}."""
    plugins: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = []
    seen: set[str] = set()
    for root in plugin_roots(paths):
        if not root.is_dir():
            continue
        for entry in sorted(root.iterdir()):
            if not entry.is_dir() or not (entry / "plugin.json").is_file():
                continue
            try:
                manifest = load_manifest(entry)
            except PluginValidationError as exc:
                errors.append({"dir": str(entry), "message": str(exc)})
                continue
            name = manifest["name"]
            if name in seen:
                errors.append({"dir": str(entry), "message": f"duplicate plugin name '{name}' (first one wins)"})
                continue
            seen.add(name)
            plugins.append(
                {
                    "name": name,
                    "version": str(manifest.get("version", "")),
                    "description": manifest["description"],
                    "dir": str(entry),
                    "script": str(manifest["script"]),
                    "input_schema": manifest.get("input_schema"),
                    "interpreter": manifest.get("interpreter"),
                    "timeout_ms": manifest.get("timeout_ms"),
                }
            )
    return {"plugins": plugins, "errors": errors}


def resolve_plugin(name_or_dir: str, paths: str | None = None) -> dict[str, Any]:
    """Resolve a plugin by unique name or by directory path."""
    discovered = discover_plugins(paths)
    candidate = Path(name_or_dir)
    if candidate.is_dir():
        manifest = load_manifest(candidate)
        matches = [p for p in discovered["plugins"] if p["name"] == manifest["name"]]
        entry = matches[0] if matches else {
            "name": manifest["name"],
            "version": str(manifest.get("version", "")),
            "description": manifest["description"],
            "dir": str(candidate.resolve()),
            "script": str(manifest["script"]),
            "input_schema": manifest.get("input_schema"),
            "interpreter": manifest.get("interpreter"),
            "timeout_ms": manifest.get("timeout_ms"),
        }
        return entry
    matches = [p for p in discovered["plugins"] if p["name"] == name_or_dir]
    if not matches:
        raise PluginValidationError(f"plugin '{name_or_dir}' not found in {[str(r) for r in plugin_roots(paths)]}")
    if len(matches) > 1:
        raise PluginValidationError(f"plugin name '{name_or_dir}' is ambiguous; pass a directory path")
    return matches[0]


def run_plugin(
    name_or_dir: str,
    params: dict[str, Any] | None = None,
    context: dict[str, Any] | None = None,
    *,
    paths: str | None = None,
) -> dict[str, Any]:
    """Run one plugin as a subprocess with the stdin JSON contract."""
    entry = resolve_plugin(name_or_dir, paths)
    script = Path(entry["dir"]) / entry["script"]
    interpreter = entry.get("interpreter") or sys.executable
    timeout_ms = entry.get("timeout_ms") or 120_000
    payload = {"context": context or {}, "params": params or {}}
    try:
        proc = subprocess.run(
            [str(interpreter), str(script)],
            input=json.dumps(payload),
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=timeout_ms / 1000.0,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        return {"success": False, "plugin": entry["name"], "reason": f"plugin timed out after {timeout_ms} ms", "stderr": str(exc)}
    if proc.returncode != 0:
        return {
            "success": False,
            "plugin": entry["name"],
            "reason": f"plugin exited with code {proc.returncode}",
            "stderr": proc.stderr.strip()[:2000],
        }
    try:
        result = json.loads(proc.stdout)
    except json.JSONDecodeError as exc:
        return {
            "success": False,
            "plugin": entry["name"],
            "reason": f"plugin stdout is not JSON: {exc}",
            "stdout": proc.stdout[:2000],
        }
    result.setdefault("plugin", entry["name"])
    result.setdefault("stderr", proc.stderr.strip()[:2000])
    return result
