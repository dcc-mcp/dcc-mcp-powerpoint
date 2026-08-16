# Learnings — powerpoint-deck

Dated lessons applied to this skill (learnings loop; append, never edit
history).

## 2026-08-16 — showcase deck optimization + ecosystem research

- Web research across document-pptx (vasilyu1983/AI-Agents-public),
  IBM/chuk-mcp-pptx, and Tencent/Youtu-agent ppt_gen confirmed the core
  design: template-first with layout resolution by name, schema-driven
  content (their YAML type_map + field constraints ≙ our Deck IR +
  validation), and a component registry with LLM-friendly schemas (≙
  office-tools registry).
- Adopted: alt text on every picture (a11y as authoring), inventory tool
  before repair, decision rules + known-limits in SKILL.md, and this
  learnings loop.
- Gotcha found live: a site asset named .png was actually WEBP content —
  python-pptx rejects WEBP; always verify magic bytes, not extensions.
- Showcase imagery now embeds normalized 480x270 cards (16:9, alpha-padded);
  image_grid layout renders slide.images as captioned cards with explicit
  missing_asset notes.

## 2026-08-16 — OfficeCLI (iOfficeAI) 技术调研与边界决策

调研范围：README + 官方 SKILL.md + 实测 CLI（v1.0.144，win-x64）。

### 接口面（供自研参考）

- 单二进制、无 Office 依赖；内置 HTML 渲染引擎（docx/xlsx/pptx → HTML/PNG），
  形成 render → look → fix 闭环（给 agent 眼睛）
- 稳定路径寻址：`/slide[1]/shape[@id=N]`、`/body/table[1]`、`/Sheet1/A1`；
  `get <path> --json` 结构化回读；`query` 支持 CSS-like 选择器
- `view` 多模式：outline / stats / issues / text / annotated / html /
  svg / screenshot / pdf / forms；issues 分 format/content/structure 桶，
  pptx 子类型含 low_contrast、broken_part_ref、notes_unresolved_rid
- issues 质量高：能给出具体修复提示（实测抓到我们 timeline 圆形数字
  "12pt 需要 16pt" 的真实溢出，suggest.height=1.4cm）
- L1 read → L2 DOM edit → L3 raw XML 分层策略；resident 会话 + 只在非
  officecli 边界 flush；MCP 采用单 `command` 字符串透传（与我们任务级
  结构化工具是两种取舍）
- per-script 字体槽（lang.latin/ea/cs）——已自研采纳 a:cs 槽

### 边界决策（2026-08-16）

- 核心功能自研，不引入外部二进制依赖；OfficeCLI 仅作接口设计参考
- 自研 roadmap（需要时实现）：
  1. 无 Office 的预览路径（HTML/PNG 渲染引擎，render→look→fix 闭环）
  2. issues 分析器：路径寻址 + format/content/structure 分类 + 具体修复提示
  3. inventory 输出采用 `/slide[i]/shape[j]` 路径寻址
  4. 低对比度（low_contrast）检查并入视觉 QA
