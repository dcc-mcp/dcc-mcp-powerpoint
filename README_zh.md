# dcc-mcp-PowerPoint

DCC-MCP 生态的 PowerPoint 适配器 —— 作为
[dcc-mcp-office](https://github.com/dcc-mcp/dcc-mcp-office) 之上的
**应用适配层**：从结构化内容生成 Deck、Slide 编排、DCC 渲染产物评审 Deck、
批量转 PDF、批量替换文字、逐页预览与视觉校验。

**状态：** M0 骨架。协议/IR 与 C# `office-host` 骨架在
`dcc-mcp-office` 落地；M1（COM MVP）就绪后本仓库在其上实现 PowerPoint
应用语义。

## 设计（两层 Sidecar）

```text
Agent → dcc-mcp-gateway (Rust, dcc-mcp-core)
          → dcc-mcp-server sidecar  (生命周期：RFC #998)
              → office-host.exe --app=powerpoint  (C#, dcc-mcp-office)
                  → POWERPNT.EXE (COM, 当前登录用户会话)
```

## Skill Pack

- `powerpoint-deck` — 从结构化内容生成可编辑 Deck（模板 → 语义布局 →
  IR → Open XML → COM 终稿 → 预览 → 校验循环）。
- `powerpoint-review` — `dcc.review-deck-from-renders`：从
  Maya/Houdini/Blender/Unreal 渲染与 Shot/Asset 元数据生成评审 Deck。

详见 [AGENTS.md](./AGENTS.md)。

## 许可

MIT
