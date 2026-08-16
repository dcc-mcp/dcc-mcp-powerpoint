# ADR 003 — PyOxidizer 独立可执行文件分发

- **Status**: Accepted
- **Date**: 2026-08-16
- **Related**: dcc-mcp-photoshop pyoxidizer.bzl 先例、CONTRIBUTING 依赖策略

## Context

依赖策略要求运行时零第三方依赖，且重型 Office 能力由我们的 C# 宿主提供。
端用户分发需要一条不依赖 pip/Office 环境的分发路径。dcc-mcp-photoshop 已用
PyOxidizer 产出独立可执行文件（学习范本）。

## Decision

- 采用 PyOxidizer（构建期工具，非运行时依赖）打包
  `dcc-mcp-powerpoint.exe`：
  - `resources_location = "filesystem-relative:lib"`（photoshop 关键经验：
    dcc-mcp-core 与技能包依赖 `__file__`，内存模式会破坏它）
  - 标准文件系统 importer（`oxidized_importer = False`）、
    `parse_argv = False`
  - `run_module = "dcc_mcp_powerpoint._standalone_entry"`：双入口
    （CLI 命令 + 技能脚本透传，与 photoshop 相同模式）
- **C# 宿主作为旁挂资源打进 lib/**：`vendor/lib/dcc-office-host.exe`
  → FileManifest.add_path 去前缀后落位 lib/；宿主解析顺序
  DCC_OFFICE_HOST → $ORIGIN/lib → PATH
- 宿主由 `tools/build_binary.py` 先用 `dotnet publish`（self-contained
  win-x64）构建，再跑 pyoxidizer
- 首期仅 Windows x64（Office/COM 域）；Linux/macOS 的 Office-less 路径
  待宿主多平台发布后开放

## Consequences

- 单目录分发：exe + lib/（Python 资源 + 宿主），压缩后一个 zip
- 运行时无 pip、无 Python 安装、无 NuGet 依赖
- COM 渲染暂由 Python render_deck 承担（需要桌面 Office）；宿主接管
  COM 后独立包也随迁移获得完整能力
- CI 后续加 release job（矩阵 build + 冒烟 + 上传产物），对齐 photoshop
  的 release.yml 模式
