# sci-flowchart-vba

[🇬🇧 English](#english) · [🇨🇳 中文](#zh)

---

<h2 id="english">🇬🇧 English</h2>

A WorkBuddy skill that turns SCI-style flowchart figures (and paper text) into
**editable PowerPoint files** — a real `.pptx` plus the VBA source that can
rebuild it at any time.

### What it does

Given a flowchart image, the skill produces VBA code that draws native,
editable PowerPoint shapes (not a picture). Paste the `.bas` modules into
PowerPoint's VBA editor and run `BuildFlowchart`, or just open the generated
`.pptx` — the shapes are already drawn and fully editable.

### Two modes

- **Mode 1 — plan + generate**: supply a style template figure + a paper
  paragraph. The skill borrows the template's visual style (palette, rounded
  corners, stroke, font) and the paragraph's content to author a new flowchart.
- **Mode 2 — replicate**: supply an existing flowchart image; the skill rebuilds
  its logical structure in SCI style.

### Why vision-direct

The old pipeline rasterized the figure to SVG first, then parsed it through
three scripts. That intermediate step lost shape semantics, exact colors, and
text. This skill drops SVG entirely: the image goes straight to a multimodal
model, which emits the *content module* (`modFlow_Content.bas`). A fixed engine
(`assets/modFlow_Engine.bas`) owns all the hard parts — coordinate mapping,
shape creation, connectors, text, idempotent cleanup, and the entry point.

### Layout

```
sci-flowchart-vba/
  SKILL.md              # skill manifest + full contract for the code model
  assets/
    modFlow_Engine.bas  # fixed engine (do not rewrite)
  references/
    example_content.bas # runnable content-module example
    vba-module-map.md   # engine API table + content-module contract
    review_rubric.md    # two-round review standards (style + rendered geometry)
    review_examples/    # origin.jpg (good) / failed.png (bad) sample pair
  scripts/
    build_ppt.py        # final step: .bas -> real .pptx (COM, replay fallback)
    lint_vba.py         # pre-delivery structural self-check
    review_render.py    # review round 2: overflow / dangling / overlap checks
    render_preview.py   # replay .bas -> preview_compare.png
    extract_colors.py / extract_geometry.py / compare_text.py
```

### Quick start (as a user)

1. Get the `.pptx` (the skill already generates it via `build_ppt.py`).
2. Or rebuild yourself: in PowerPoint press `Alt+F11`, right-click the project,
   **Import File** each `.bas` (engine + content), then `F5` → `BuildFlowchart`.
3. All shapes are editable. Run `RemoveFlowchart` to clear and redraw.

### Contract highlights (for the code model)

- Content module must declare 9 geometry constants and a single
  `Public Sub DrawAll(sld As Slide)`.
- Colors are `RGB(r,g,b)` literals; no `vbRed` etc.
- `.bas` files are **UTF-8 encoded + CRLF** — node text and comments may be
  Chinese (keep the source language; for a Chinese figure write Chinese directly).
  UTF-8 is what lets Chinese survive into the `.pptx`.
- Split into `modFlow_Content2.bas` when nodes > 18 or a module > 300 lines.
- Math (radicals, fractions, sub/superscripts) uses `AddFormula sld, id, x, y, w, h, "<LaTeX>", "<linear fallback>", fontsize, color`
  producing a **native editable Office equation** (LaTeX->OMML, both paths); simple variables stay as italic Times labels.
- **Two review rounds are mandatory before delivery**: round 1 checks style
  fidelity to the source figure (palette saturation, font hierarchy, stroke
  weights); round 2 checks the rendered `.pptx` (`review_render.py`) for
  oversized text, dangling connectors, and misplaced arrows.

### License

MIT — do whatever you like; attribution appreciated.

---

<h2 id="zh">🇨🇳 中文</h2>

一个 WorkBuddy 技能，把 SCI 风格的流程图插图（以及论文正文）转换成
**可编辑的 PowerPoint 文件**——一份真正能打开的 `.pptx`，外加可随时重建它的
VBA 源码。

### 它做什么

给定一张流程图图片，技能会生成绘制「原生、可编辑」PowerPoint 形状的 VBA 代码
（不是图片）。把 `.bas` 模块导入 PowerPoint 的 VBA 编辑器并运行 `BuildFlowchart`，
或者直接打开生成好的 `.pptx`——形状已经画好，且全部可编辑。

### 两种模式

- **模式① 规划+生成**：提供一张风格模版图 + 论文某段文字。技能从模版图借用
  视觉风格（配色、圆角、描边、字体），从论文取内容，规划架构后出一张新流程图。
- **模式② 复刻**：提供一张已有的流程图图片；技能按 SCI 风格重建它的逻辑结构。

### 为什么用 vision-direct

旧方案先把图矢量化成 SVG，再经三个脚本解析。这一步「先转 SVG」会丢掉形状语义、
精确配色和文字。本技能完全去掉 SVG：图直接交给多模态大模型，由它产出*内容模块*
（`modFlow_Content.bas`）。固定引擎（`assets/modFlow_Engine.bas`）负责所有硬骨头——
坐标映射、形状创建、连线、文字、幂等清理和入口过程。

### 目录结构

```
sci-flowchart-vba/
  SKILL.md              # 技能清单 + 给代码模型的完整契约
  assets/
    modFlow_Engine.bas  # 固定引擎（勿改写）
  references/
    example_content.bas # 可运行的内容模块示例
    vba-module-map.md   # 引擎 API 全表 + 内容模块契约
    review_rubric.md    # 两轮评审标准（风格一致性 + 渲染几何）
    review_examples/    # 正/负样本对照（origin.jpg / failed.png）
  scripts/
    build_ppt.py        # 最后一步：.bas -> 真实 .pptx（COM 优先，回放兜底）
    lint_vba.py         # 交付前的结构自检
    review_render.py    # 评审二：文字溢出 / 悬空连线 / 重叠检查
    render_preview.py   # 回放 .bas 渲染 preview_compare.png
    extract_colors.py / extract_geometry.py / compare_text.py
```

### 快速开始（作为使用者）

1. 拿到 `.pptx`（技能已通过 `build_ppt.py` 生成）。
2. 或自己重建：PowerPoint 里按 `Alt+F11`，右键工程 → **Import File** 导入每个
   `.bas`（引擎 + 内容），然后 `F5` → `BuildFlowchart`。
3. 所有形状都可编辑。运行 `RemoveFlowchart` 清空并重画。

### 出码契约要点（给代码模型）

- 内容模块必须声明 9 个几何常量，并定义唯一的 `Public Sub DrawAll(sld As Slide)`。
- 颜色用 `RGB(r,g,b)` 字面量；不要 `vbRed` 之类。
- `.bas` 文件必须**UTF-8 编码 + CRLF 换行**——节点文字与注释可用中文（源图是中文时直接写中文），UTF-8 才能让中文正常落地。
- 节点 > 18 或单模块 > 300 行时，拆到 `modFlow_Content2.bas`。
- 数学式（根号 / 分式 / 上下标）用 `AddFormula sld, id, x, y, w, h, "<LaTeX>", "<线性fallback>", fontsize, color` 
  直出 **原生可编辑 Office 公式**（LaTeX→OMML，双路径可编辑）；简单变量用斜体 Times 即可。
- **交付前必须过两轮评审**：评审一对照源图查风格一致性（配色饱和度、字号层级、
  线框粗细）；评审二用 `review_render.py` 检查最终 `.pptx`（字体是否过大、
  连线是否悬空、箭头位置是否正确）。

### 许可证

MIT——可随意使用，注明出处更佳。
