---
name: powerpoint-plugins
description: >-
  Discover and invoke user-installed PowerPoint plugins. A plugin is a
  directory with a plugin.json manifest (name/version/description/script/
  optional input_schema) plus a script that reads stdin JSON
  {"context": ..., "params": ...} and prints one JSON result. Plugins are
  discovered from DCC_POWERPOINT_PLUGIN_PATH and
  ~/.dcc-mcp/powerpoint/plugins, validated strictly (slug name, script
  inside the plugin dir, timeout cap), and run as isolated subprocesses.
  Use to extend deck workflows with project-specific transforms without
  touching the adapter.
license: MIT
allowed-tools: Bash Read
metadata:
  dcc-mcp:
    dcc: powerpoint
    layer: domain
    stage: authoring
    version: 0.1.0
    tags:
      - powerpoint
      - plugins
      - extension
      - registry
      - pptx
    search-hint: >-
      powerpoint plugin, invoke plugin, run plugin, plugin list, extension,
      custom pptx transform, plugin registry
    tools: tools.yaml
---

# powerpoint-plugins (Authoring stage)

Extension point for project-specific deck transforms. The adapter never
fetches or installs plugin code: plugins live in explicit local directories
and run as plain subprocesses under the skill's execution contract — the
same stdin-JSON → stdout-JSON convention as skill scripts themselves.

## Plugin layout

```
plugins/
  my-splitter/
    plugin.json   # {"name", "version", "description", "script", ...}
    splitter.py   # reads stdin JSON, prints one JSON result
```

## plugin.json contract

| Key | Required | Meaning |
|---|---|---|
| name | yes | slug [a-z0-9-_], unique across discovery roots |
| version | yes | free-form version string |
| description | yes | one line, surfaced in tool descriptions |
| script | yes | script path relative to the plugin dir (may not escape it) |
| interpreter | no | default: the running Python |
| input_schema | no | JSON Schema for params (surfaced to agents) |
| timeout_ms | no | 1..600000, default 120000 |

The script receives stdin JSON `{"context": {...}, "params": {...}}` and
must print one JSON object with at least `{"success": bool}`. stderr is
captured into the result for diagnostics.

## Discovery roots

1. `DCC_POWERPOINT_PLUGIN_PATH` (os.pathsep-separated directories)
2. `~/.dcc-mcp/powerpoint/plugins`

## Decision rules

- `plugin_list` first — discover names and schemas before running
- pass the deck via context: `{"pptx": "deck.pptx"}` plus any plugin
  needs (workdir, output dir, manifest files)
- treat plugins as user-installed code: run only what the user installed,
  never write a plugin to paper over a skill gap that belongs in the repo
- a plugin that times out or exits non-zero returns a structured failure;
  never retry blindly

## Known limits

- plugins run with the gateway process's user permissions; the registry
  adds process isolation and timeouts, not a security sandbox
- one plugin = one subprocess call; chained transforms are separate calls

## Agent-visible summary

Discovered plugins with schemas, plus the plugin's result, stderr
diagnostics and timing on run.
