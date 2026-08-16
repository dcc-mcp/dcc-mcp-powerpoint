# 能力边界 Showcase — dcc-mcp-PowerPoint

本目录是 **capability boundary** 演示：一份 17 页 Deck 走完仓库全部能力面，
再用三类生成后能力（图层 / 补丁 / 插件）二次加工。所有产物由本仓库的
技能脚本生成，可作为回归基准。

## 能力覆盖清单

| 能力面 | 覆盖方式 | 证据 |
|---|---|---|
| 11 种语义布局 | showcase IR 逐一使用 | `capability_boundary_showcase.json`（17 slides, 0 warnings） |
| 生成管线 | Deck IR → Open XML → COM 渲染 → 验证 | `output/capability-boundary/*.pptx|pdf|previews` |
| 语义图层 list/显隐/重着色/重排/归属 | 5 个 layer 脚本各出一份变体 | `output/capability-boundary/variants/` |
| 补丁引擎（7 类操作，全有或全无） | ppt-patch/1.0 文档一次应用 | `patches/showcase-v2.patch.json` → `variants/patched-v2.pptx` |
| 插件注册表（发现 + 隔离子进程执行） | deck-stats 演示插件 | `../../plugins/deck-stats/`（plugin.json + script） |
| 问题分析器 | analyze_deck 全量扫描 | 溢出/对比度/字体/alt 缺失分桶报告 |

## 复现

```bash
# 1. 生成（编译 + COM 渲染 PDF/previews + 结构验证）
python src/dcc_mcp_powerpoint/skills/powerpoint-deck/scripts/generate_deck.py \
  --input examples/capability_boundary_showcase.json \
  --out examples/output/capability-boundary

# 2. 图层演示（list / 显隐 / 重着色 / 重排 / 归属）
python src/dcc_mcp_powerpoint/skills/powerpoint-layers/scripts/layer_list.py \
  --input examples/output/capability-boundary/draft-capability-boundary-showcase.pptx
python src/dcc_mcp_powerpoint/skills/powerpoint-layers/scripts/layer_set_visibility.py \
  --input examples/output/capability-boundary/draft-capability-boundary-showcase.pptx \
  --layer decoration --visible false \
  --output examples/output/capability-boundary/variants/no-decoration.pptx
python src/dcc_mcp_powerpoint/skills/powerpoint-layers/scripts/layer_recolor.py \
  --input examples/output/capability-boundary/draft-capability-boundary-showcase.pptx \
  --layer accent --fill 7FC8A9 \
  --output examples/output/capability-boundary/variants/green-accent.pptx
python src/dcc_mcp_powerpoint/skills/powerpoint-layers/scripts/layer_reorder.py \
  --input examples/output/capability-boundary/draft-capability-boundary-showcase.pptx \
  --layer footer --position front \
  --output examples/output/capability-boundary/variants/footer-front.pptx
python src/dcc_mcp_powerpoint/skills/powerpoint-layers/scripts/layer_assign.py \
  --input examples/output/capability-boundary/draft-capability-boundary-showcase.pptx \
  --layer callout --select '{"slides": [3], "shape": "TextBox 5"}' \
  --output examples/output/capability-boundary/variants/assigned-callout.pptx

# 3. 补丁引擎（7 操作：set_text/set_notes/set_shape_fill/set_shape_visible/
#    add_slide/move_slide/assign_layer，任一失败即整体回滚）
python src/dcc_mcp_powerpoint/skills/powerpoint-edit/scripts/edit_deck.py \
  --input examples/output/capability-boundary/draft-capability-boundary-showcase.pptx \
  --patch examples/output/capability-boundary/patches/showcase-v2.patch.json \
  --output examples/output/capability-boundary/variants/patched-v2.pptx

# 4. 插件（发现 + 执行）
python src/dcc_mcp_powerpoint/skills/powerpoint-plugins/scripts/plugin_list.py \
  --paths examples/plugins
python src/dcc_mcp_powerpoint/skills/powerpoint-plugins/scripts/plugin_run.py \
  --name deck-stats --paths examples/plugins \
  --pptx examples/output/capability-boundary/draft-capability-boundary-showcase.pptx
```

## 已知基线

`analyze_deck` 对结构桶输出文本溢出提示（文本框按行高估算，多行文本会
提示建议高度）；两份 deck（framework-intro 与 capability-boundary）均为
~109 条结构提示，属当前分析器基线，不是本 showcase 回归。
