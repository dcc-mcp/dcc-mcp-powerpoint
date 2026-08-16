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

## 2026-08-16 — 依赖边界：自研 + 自有 C# 宿主

决策：dcc-mcp-PowerPoint 运行时零第三方依赖；重型能力（Open XML 读写、
COM 渲染）由我们自己的 dcc-mcp-office C# 宿主提供，Python 走 office-rpc
契约调用（先 stdin/stdout JSON-RPC，后按方案 §12 换命名管道）。

- Python 运行时只用 stdlib（zipfile/xml.etree/json/dataclasses/subprocess/ctypes）
- python-pptx/Pillow/pywin32 降级为 dev/test 专用（测试 oracle、夹具）
- analyze/inventory 的 pptx 读取层迁移到 C# 宿主（当前例外，已记录）
- C# 宿主零非微软 NuGet：System.IO.Packaging + LINQ to XML（net8.0-windows
  内置）+ BCL COM 互操作
- 第三方资产（模板/字体/素材）仅经许可核验 + 自托管后入库，构建/运行时
  不从外部源拉取

## 2026-08-18 — 编辑 / 智能图层 / 插件三技能迭代（feat/smart-layers-plugin-skills）

- 新增智能图层系统（layers.py）：图层成员关系记录在形状原生名称的
  `::layer=<name>` 标签里（cNvPr/@name），零 sidecar、PowerPoint 往返无损；
  编译器现在给每个形状打上六大内置图层 background/decoration/header/
  content/accent/footer —— 生成即图层化。图层操作直改 spTree：hidden 属性
  （显隐）、子元素重排（z-order front/back/above/below）、显式 solid 填充/
  线条/run 颜色换色（主题继承色诚实跳过，绝不猜测）。
- 新增 ppt-patch/1.0 补丁引擎（edits.py）：结构化、坐标无关的 deck 修改。
  selector = slide/layer/shape/id/index 的组合；操作覆盖文本、备注、换图、
  填充/描边、显隐、删除形状、增删移幻灯片（新幻灯片复用编译器语义布局
  注册表）、图层打标。事务语义：任何操作失败则在保存前中止 —— 输入文件
  不会被改坏。默认输出 <stem>-patched.pptx。
- 新增插件注册表（plugins.py，纯 stdlib）：插件 = 目录 + plugin.json 清单
  （slug 名、script 不得逃逸插件目录、timeout 上限 600s）；发现根仅为
  DCC_POWERPOINT_PLUGIN_PATH + ~/.dcc-mcp/powerpoint/plugins，绝不自动
  下载/安装；执行 = 子进程 + stdin JSON 契约（与 skill 脚本同款约定），
  stderr 回传诊断。
- 三个新技能包按现代工具契约发布（execution/job_strategy/affinity/
  enforce_thread_affinity/annotations/call_examples）：powerpoint-layers、
  powerpoint-edit（depends: powerpoint-deck）、powerpoint-plugins。
- powerpoint-deck 契约现代化：全部工具补 call_examples 等字段，SKILL.md
  升 0.2.0 并挂 references/RECIPES.md（每个语义布局一份可复制的 IR 样例）；
  powerpoint-review 同步现代化。
- 边界：edits/layers 与 compiler 一样用 python-pptx（dev 依赖、opt-in 导入，
  C# 宿主迁移仍按 learnings 既定路线）；插件执行在本沙箱因命名管道限制
  无法本地冒烟，以 monkeypatch 单测覆盖 subprocess 契约。

