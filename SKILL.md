---
name: sci-flowchart-vba
description: 把 SCI 论文风格的流程图插图转化为可编辑 PowerPoint 文件（.pptx + VBA 源码）。当用户提供流程图模版图片/模版库文件夹 + 论文某部分文字（模式①：规划流程图架构并按其风格出图），或提供一张已有的流程图图片（模式②：复刻该流程图），或明确提到"流程图转 PPT""流程图转 VBA""SCI 风格流程图""仿照这张图做流程图""把这张流程图复刻成可编辑 PPT"时，必须使用本 skill。最终交付物是**已画好形状、可直接打开编辑的 .pptx**，同时附带分模块的 .bas 源码（导入 PowerPoint 后运行 BuildFlowchart 可随时重建）。
agent_created: true
---

# SCI 风格流程图 → PowerPoint VBA（vision-direct 版）

把 SCI 论文插图里那种"扁平配色 + 细描边 + 圆角卡片 + 无衬线标签"的流程图，
转成**纯 VBA 代码**。用户把 `.bas` 导入 PowerPoint 后运行 `BuildFlowchart`，
得到的是**原生可编辑形状**（不是图片），可继续改文字、调颜色、挪位置。

## 核心变化（为什么这么改）

旧版先把图矢量化成 SVG，再经 `sci_flow_parse.py → layout_analyze.py →
build_vba.py` 三脚本管线生成 VBA。**这一步"先转 SVG"损失太大**：多模态
大模型本身就能看图，矢量化的中间环节反而丢掉了形状语义、精确配色和文字。
新版**完全去掉 SVG 矢量化与三脚本管线**：

- **图直接交给多模态大模型**；模型识图后直接产出 VBA 内容模块。
- 坐标映射、形状创建、连线、文字、幂等清理、入口过程，全部由**固定引擎**
  `assets/modFlow_Engine.bas` 保证正确——模型只负责它擅长的事：看懂图、
  规划结构、写出节点/连线清单。
- 不再有"像素级复刻"，而是产出**逻辑正确、风格统一、干净可编辑**的流程图。

## 两条路径

| 用户给了什么 | 模式 | 做法 |
|---|---|---|
| 一张模版图（或模版库）+ 论文某段 / 要画的内容 | **模式① 规划+生成** | 模型从模版图取"风格"（配色/圆角/描边/字体），从论文取"内容"，规划架构后出 VBA |
| 一张已有的流程图图片，要求"复刻/重画" | **模式② 复刻** | 模型看图重建逻辑结构，按 SCI 风格出 VBA |

> 若用户给的是 AI/Inkscape 导出的 SVG：导出成 PNG 后当图片走上面的路即可，
> 不要再走旧的 SVG 解析。

## 技能运行步骤（agent 照做）

1. **收输入**：模式①要模版图 + 论文段落；模式②要流程图图片。
2. **给大模型下指令**：把图片（模式①还要模版图 + 论文文字）连同下面的
   **「给大模型的出码契约」**一起发给多模态模型，要求它只回 `modFlow_Content.bas`
   的源码（可拆分多个模块）。
3. **拼装交付物**：把 `assets/modFlow_Engine.bas` **原样复制**进输出目录，
   写入大模型产出的 `modFlow_Content.bas`（及拆分的 `modFlow_Content2.bas` …）。
4. **自检**（见末尾清单）。
5. **生成 PPT**：用 `scripts/build_ppt.py <输出目录>` 把 `.bas` **自动灌进一个
   真正的 `.pptx` 并直接执行出图**（详见下节「最后一步：自动产出 PPT」）。
6. **两轮评审**（详见「两轮评审」章节）：评审一过风格一致性 → 评审二过
   渲染几何检查；任一轮不过就修 `.bas` 重跑，**通过后才准交付**。
7. **交付**：`.pptx` 优先，附上全部 `.bas` + 简明导入说明。

---

## 给大模型的出码契约（把这一段连同图片一起发给模型）

你是一个"看图写 PowerPoint VBA"的专家。用户给你一张流程图（模式①还附带
风格模版图与论文文字）。请直接产出 VBA 内容模块 `modFlow_Content.bas` 的源码，
**不要**先转 SVG、不要写解释、只输出代码。

### 你拥有的固定引擎（勿重写，直接调用）

- `AddNode sld, id, kind, x, y, w, h, fillC, lineC, lineW, dash, [text], [fontPt], [bold], [italic], [fontC], [cornerRatio], [fontName], [adj2]`
  建一个节点。坐标 (x,y) 为左上角、单位 px（指你下面的布局画布，不是图片真实像素）。
  `cornerRatio` 对圆角矩形是圆角比例、对块箭头是箭头相对高度；`adj2` 只给块箭头用（箭头相对长度）。
- `AddFormula sld, id, x, y, w, h, "<LaTeX>", "<linear fallback>", fontPt, fontC`
  在 (x,y,w,h) 里放一个**原生 Office 公式**（LaTeX → OMML），居中、Cambria Math。
  适合分式 / 根号 / 上下标 / 求和号这类"文本框摆不出来"的数学式；简单变量直接用
  AddNode 的斜体 Times 即可，不必上 AddFormula。
  - `<LaTeX>` 用 ASCII 转义写（`\times` `\sqrt[3]{...}` `\frac{a}{b}` 等），字符串里
    出现单个反斜杠即可，**不要双写**；VBA 里反斜杠不是转义字符。
  - `<linear fallback>` 是给"VBA 手动按 F5"备用的线性版（COM 先画它，build_ppt.py
    随后把保存的 .pptx 后处理成真 OMML，两条路最终看到的都是公式对象）。
  - 依赖：replay/后处理路径需要 `pip install latex2mathml mathml2omml`（venv 已装）。
  - 预览渲染用 matplotlib mathtext 画公式，与 PowerPoint 的 Cambria Math 成像略有
    字形差异，但版式（根号/分数结构）一致。
- `AddPath sld, id, "x1,y1;x2,y2;x3,y3", lineC, lineW, dash, arrowEnd`
  画一条折线连线（点串用分号分隔、逗号分隔 xy，单位 px）；`arrowEnd=True` 时末端带箭头。
- 颜色用 `RGB(r,g,b)`；填充/描边传 `-1` 表示无。`LINE_SOLID=1 / LINE_DASH=4 / LINE_DOT=2 / LINE_DASHDOT=5`。
- 形状 kind 字符串：`rect` `round_rect` `oval` `diamond` `parallelogram`
  `trapezoid` `pentagon` `hexagon` `stadium`(药丸/起止) `cylinder` `document`
  `right_arrow` `left_arrow` `up_arrow` `down_arrow` `chevron`（块箭头/连接箭头）。
  圆角矩形用 `round_rect` 并给 `cornerRatio`（0~0.5，如 0.12）；`stadium` 自动大圆角。

### 模块必须长这样（顺序固定）

```vba
Attribute VB_Name = "modFlow_Content"
Option Explicit
Option Base 0

' 1) 几何常量（引擎引用，缺一会编译失败）
Public Const CANVAS_W_PX As Single = 1200     ' 你的布局画布宽
Public Const CANVAS_H_PX As Single = 675      ' 你的布局画布高（与图同宽高比）
Public Const PX_TO_IN As Single = 0.0108      ' = min(availW/CANVAS_W, availH/CANVAS_H)
Public Const OFFSET_X_IN As Single = 0.62     ' 居中偏移
Public Const OFFSET_Y_IN As Single = 0.35
Public Const SLIDE_W_IN As Single = 13.333    ' 16:9；若图更高见下方"选版面"
Public Const SLIDE_H_IN As Single = 7.5
Public Const CUSTOM_SIZE As Boolean = False   ' 非 16:9 时改 True 并填上面 SLIDE_*
' 字体：与源图字形一致。西文衬线 Times New Roman / 西文无衬线 Arial；
' 中文黑体/无衬线 Microsoft YaHei、中文宋体/衬线 SimSun、粗黑 SimHei。
' 预览与评审脚本会按 FONT_NAME 自动匹配对应 CJK 字体文件（YaHei/SimSun/SimHei）。
Public Const FONT_NAME As String = "Arial"

' 2) 调色板（从图取色，一律 RGB 字面量）
Public Const INK As Long = RGB(31, 31, 31)
Public Const BG As Long = RGB(255, 255, 255)
Public Const ACCENT As Long = RGB(68, 114, 196)
Public Const ALT As Long = RGB(226, 237, 250)
Public Const LINE As Long = RGB(120, 120, 120)

' 3) 唯一入口（BuildFlowchart 会调用它）
Public Sub DrawAll(ByVal sld As Slide)
    DrawContent1 sld
    ' 若拆分：DrawContent2 sld  ...
End Sub

' 4) 内容：节点 + 连线
Public Sub DrawContent1(ByVal sld As Slide)
    AddNode sld, "start", "stadium", 540, 20, 120, 50, ACCENT, -1, 1, LINE_SOLID, _
                "Start", 14, True, False, INK
    AddNode sld, "a", "round_rect", 520, 130, 160, 60, BG, LINE, 1, LINE_SOLID, _
                "Collect data", 12, False, False, INK, 0.12
    ' 中文流程图：节点文字直接写中文即可（.bas 用 UTF-8 编码）
    AddNode sld, "cn", "round_rect", 700, 130, 160, 60, BG, LINE, 1, LINE_SOLID, _
                "数据收集", 12, False, False, INK, 0.12
    AddPath sld, "e1", "600,70;600,130", LINE, 1, LINE_SOLID, True
End Sub
```

### 几何常量怎么算（务必算对，否则图会跑偏/溢出）

1. **选版面**：图比宽 → 16:9（`SLIDE_W_IN=13.333, SLIDE_H_IN=7.5`，`CUSTOM_SIZE=False`）；
   接近方 → 4:3（`10, 7.5`）；明显更高 → 竖版（`7.5, 10`）或方形（`8.5, 8.5`），
   此时 `CUSTOM_SIZE=True`。
2. 设 `CANVAS_W_PX / CANVAS_H_PX` 为一个与图**同宽高比**的布局画布（如 16:9 用 1200×675）。
3. `margin = 0.35`（英寸）。`availW = SLIDE_W - 2*margin`，`availH = SLIDE_H - 2*margin`。
4. `PX_TO_IN = min(availW/CANVAS_W, availH/CANVAS_H)`。
5. `OFFSET_X = margin + (availW - CANVAS_W*PX_TO_IN)/2`，`OFFSET_Y` 同理。
6. 节点坐标就在这个 `CANVAS_W×CANVAS_H` 画布里用整数摆（像上面示例那样）。

### 硬约束

- **文字保留源语言（中文图就写中文）**：节点文字、注释都可以直接用中文，
  源图是中文流程图时**直接保留原中文**，不要翻译成英文 / 缩写 / 拼音。
  `.bas` 统一用 **UTF-8 编码 + CRLF 换行**（不要存成 Latin-1 / ANSI）。
  - 回放路径（`build_ppt.py` 无 Office 时）与 COM 路径都按 UTF-8 读取，
    最终 `.pptx` 里的中文与源图一致。
  - 若要在中文 Windows 上手动用 VBE「Import File」导入，导入前把 `.bas`
    另存为 ANSI（GBK，记事本"另存为 → 编码：ANSI"）即可；用 `build_ppt.py`
    直接出图则无需此步。
- **强制换行**：需要"两行文字"时写 `"Revenue" & vbLf & "Optimization"`，
  比让引擎自动折行更接近原图的排版。
- **坐标 / 颜色 / 形状 / 文字全部来自识图**，禁止凭空猜测；
  拿不准的颜色就近取图上的实际色，拿不准的位置用整齐的网格对齐。
- **字号** `fontPt` 用 9~18 之间的合理值（小框用小字），统一用 `SetLabel` 的
  自动缩放兜底，不会溢出。**层级比例要和原图一致**（主标题 > 横幅 > 节点
  标签 > 明细），宁可测（`compare_text.py`），不要一律 12pt。
  中文源图校准字号时加 `--cjk`：单行中文文字带**高度 ≈ 字号**，
  用高度比反推比宽度比可靠（宽度还受字数影响）。
- **线宽 / 配色从原图测得**：节点细描边 1pt 级；强调色保持原图饱和度、
  浅色衬底保持原图深浅——评审一会逐项对照，褪色/过淡/虚线消失都打回。
- **拆模块**：节点 > 18 或单模块 > 300 行时，把 `DrawContentN` 拆到
  `modFlow_Content2.bas` 等，`DrawAll` 要依次调用全部（跨模块 Public Sub 全局可见）。
- 连线端点要贴到形状边界（不要让箭头悬空）；正交/阶梯连线写成多点 `AddPath`。

### 交付要求

- 只输出 `modFlow_Content.bas` 源码（及可选的 `modFlow_Content2.bas` …），用
  ```vba … ``` 代码块包裹。不要输出引擎代码（引擎由技能提供）。
- 真实可运行示例见 `references/example_content.bas`，不确定时对齐它的写法。

---

## 最后一步：自动产出 PPT（不是只给 .bas）

**`.bas` 只是中间产物，最终交付物是 `.pptx`。** 生成并自检完 `.bas` 后，
必须再跑一次 `scripts/build_ppt.py`，把 VBA 灌进一个真实演示文稿并执行出图，
让用户拿到"打开就能看、能改"的 PPT。

```bash
python scripts/build_ppt.py <输出目录> [--out 交付.pptx] [--no-run] [--replay]
```

脚本做四件事：

1. 按**固定顺序**把 `.bas` 写进 `modFlow_Engine` + `modFlow_Content*` 模块
   （引擎先，内容后；顺序错了 `BuildFlowchart` 找不到 `DrawAll`）。
2. 建一个空白 16:9 演示文稿，按内容模块的 `CUSTOM_SIZE / SLIDE_W_IN / SLIDE_H_IN`
   设置页面尺寸。
3. 直接调用 `BuildFlowchart` 在幻灯片上**真的画出形状**（不是只存代码）。
4. 另存为 `.pptx`（同时保留 `.pptm`），保持全部形状可编辑。

结束后脚本会打印用了哪条路径、形状总数与文件大小；**形状数为 0 视为失败**
（说明内容没画出来），此时退出码非 0，必须回查 `.bas`。

### 依赖与降级

- 首选 **Windows + 已安装 PowerPoint 的 COM 自动化**（`win32com.client` /
  `pywin32`）。走这条路时，脚本导入模块后**真的执行 `BuildFlowchart`**，
  结果与用户手动按 `F5` 完全一致。
- 若本机无 PowerPoint，脚本自动回退到内置的 **`.bas` → 形状 → `.pptx` 回放器**：
  用 `render_preview.py` 同一套语法解析 `AddNode` / `AddPath` 实参，
  解析内容模块的几何常量算出同样的 px→英寸映射，再经 `python-pptx`
  生成等价的原生形状（坐标 / 颜色 / 圆角 / 箭头一致）。没有 Office 也能交付 PPT。
- 强制指定：`--replay` 跳过 COM 直接用回放；`--no-run` 只生成带 VBA 的
  `.pptm` 模板而不执行宏（给"想自己在 PowerPoint 里按 F5"的用户）。
- 两条路都产出可编辑 `.pptx`；差别只在"是否由 PowerPoint 引擎亲自绘制"。

### 交付要求

- 输出目录最终必须同时有：`modFlow_Engine.bas`、`modFlow_Content*.bas`、
  **`*.pptx`**、预览图。
- 用 `present_files` 把 `.pptx` 放在**第一位**呈现给用户（PPT 是主交付物）。
- 汇报时说明走的是 COM 路径还是回放路径，以及本机是否装过 PowerPoint。
- `python-pptx` / `pywin32` 不在系统解释器里时，用隔离 venv：
  `C:/Users/Hello/.workbuddy/binaries/python/envs/default/Scripts/python.exe`
  （已预装 python-pptx、pywin32、Pillow）。

---

## 两轮评审（交付前必过）

`build_ppt.py` 出了 `.pptx` 不等于完工。**必须顺序过两轮评审**，标准全文见
`references/review_rubric.md`；`references/review_examples/` 里有一对
正/负样本（`origin.jpg` 风格基准原图、`failed.png` 失败案例渲染图），
评审前先对照看一遍，把失败案例踩过的坑（红色标题褪色、面板衬底过淡、
虚线框消失、文字穿底、块箭头错位、连线断开）记在脑子里。

### 评审一：学术风格与源图一致性（.bas 调色板/字号/线宽 + 预览图）

**对象**：`modFlow_Content*.bas` 的调色板与字号常量 +
`scripts/render_preview.py` 产出的 `preview_compare.png`。
**程序化互验**：`scripts/extract_colors.py`（取色）、`scripts/compare_text.py`
（按文字带暗像素宽度反推字号）、`scripts/extract_geometry.py`（取框）。

- **配色**：每个颜色常量必须来自程序化取色，禁止目测。重点盯三类漂移：
  ① 强调色饱和度（原图鲜红标题不许褪成暗粉）；② 浅色衬底深浅（面板底色
  不许淡到近乎白色）；③ 虚线容器框颜色深度（不许淡到不可见）。
- **字体**：家族与原图一致。西文：衬线 → `Times New Roman`，无衬线 → `Arial`；
  中文：黑体/无衬线 → `Microsoft YaHei`，宋体/衬线 → `SimSun`，粗黑 → `SimHei`
  （写在 `FONT_NAME`，预览/评审脚本自动映射到对应 CJK 字体文件；单个节点
  也可用 `AddNode` 的 `fontName` 可选参数覆盖）。
  字号层级保持原图比例（主标题 > 横幅 > 节点标签 > 明细），用
  `compare_text.py` 反推（中文加 `--cjk` 按高度比），禁止一律 12pt；
  加粗位置与原图一致。
- **线框**：节点细描边 1pt 级；虚线样式/深浅同原图；块箭头颜色、方向、
  长宽比同原图。
- **风格**：扁平无阴影无渐变、对齐网格、留白均匀——学术图的基本面。

不通过 → 修常量 → 重跑 `render_preview.py` → 再对照，过了才进评审二。

### 评审二：最终 PPT 渲染检查（字体是否过大？线框/箭头位置是否正确？）

**对象**：`build_ppt.py` 产出的 `flowchart.pptx`（有 PowerPoint 就导出幻灯片
PNG 看；没有就看回放的 `preview_compare.png`，二者语义等价），
外加程序化检查：

```bash
python scripts/review_render.py <输出目录>   # 对 flowchart.pptx 自动体检
```

脚本按真实字体（PIL 量宽）检查六项：**文本溢出节点（字体过大）、依赖
自动折行、悬空连线端点、连线穿模、带文字节点意外重叠、形状出界**，
并报告全图"文字宽/框宽"中位数提示字体整体偏小。退出码 0 = PASS。

- [ ] 任何文字都不得溢出所属节点框（含自动折行后仍溢出的情形）。
- [ ] 文字与节点的比例对照原图观感一致：大而满不行，小而空也不行
      （failed.png 列表框就是"字小框空"的反例）。
- [ ] 需要两行的文字显式写 `"A" & vbLf & "B"`，不许依赖自动折行。
- [ ] 连线端点全部贴在节点边界上，无悬空；折线不穿无关节点内部。
- [ ] 块箭头在两面板衔接处、方向长度正确，不错位变形。

不通过 → 修 `.bas` → 重跑 `build_ppt.py` → 再跑 `review_render.py`。
最多迭代 3 轮；仍不过就停下向用户说明剩余问题，**不要带病交付**。

---

## 完整可运行示例

见 `references/example_content.bas`（一个"开始→收集数据→判定→训练模型→结束"
的小流程图，含 stadium / round_rect / diamond 与 4 条带箭头连线）。

## 模块职责与契约细节

见 `references/vba-module-map.md`（引擎 API 全表、形状 kind→mso 映射、内容模块契约）。

## 交付：PowerPoint 里怎么用

> 技能已通过 `scripts/build_ppt.py` 直接给出 `.pptx`（形状已画好、可直接编辑）。
> 下面这段是"想自己在 PowerPoint 里从 `.bas` 重跑一遍"时的操作。

1. 打开 PowerPoint（建议先建一个空白演示文稿）。
2. `Alt + F11` 打开 VBA 编辑器。
3. 每个 `.bas` → 右键工程 → **Import File** → 选中导入（引擎与内容都要导，顺序任意）。
4. 菜单 **Run → Run Sub**（或 `F5`），选 `BuildFlowchart` 运行。
5. 形状全部可编辑；想重来就运行 `RemoveFlowchart` 清空。
6. 若提示"宏被禁用"，在信任中心允许宏，或把文件另存为 `.pptm`。

## 自检清单（交付前过一遍）

- [ ] `modFlow_Engine.bas` 已原样附带，未被改写。
- [ ] 内容模块声明了全部 9 个几何常量（`CANVAS_W_PX … FONT_NAME`）。
- [ ] 定义了唯一 `Public Sub DrawAll(sld As Slide)`，且调用了所有 `DrawContentN`。
- [ ] 颜色全是 `RGB(r,g,b)`；无 `vbRed` 之类；节点文字可用中文（源图是中文时直接写中文）。
- [ ] 全部 `.bas` 都是 **UTF-8 编码 + CRLF 换行**（含引擎文件；中文文字/注释允许，但行尾必须是 CRLF）。
- [ ] 无空实参（形如 `AddNode …, 0.62, , 0.5` 会编译失败，中间槽位要写 `""`）。
- [ ] 连线端点贴到形状边界，无悬空箭头。
- [ ] 节点 > 18 或单模块 > 300 行时已拆到 `modFlow_Content2.bas`。
- [ ] 数学式（根号/分式/上下标）用 `AddFormula` 而非手工分层拼装；LaTeX 源写在
      字符串里（单反斜杠），fallback 给线性文本；replay/COM 都产出原生 OMML。
- [ ] **已跑 `scripts/build_ppt.py`，输出目录里有可打开的 `.pptx`，且幻灯片上确实有形状**
      （不是空板）。
- [ ] **评审一已过**：调色板/字号/线宽与原图程序化对照一致，无褪色、
      无衬底过淡、无虚线消失（标准见 `references/review_rubric.md`）。
- [ ] **评审二已过**：`scripts/review_render.py` 退出码 0，且多模态目测
      `preview_compare.png` 无文字溢出、无悬空/错位连线、块箭头位置正确。

> 项目里的 `scripts/lint_vba.py` 可以自动跑上面大部分检查：传一个输出目录即可。

## 常见问题排查

| 现象 | 原因 | 处理 |
|---|---|---|
| 编译报"变量未定义"（如 `PX_TO_IN`） | 内容模块漏了几何常量 | 补齐 9 个 `Public Const` |
| 编译报"子程序/函数未定义"（如 `DrawAll`） | 漏了 `DrawAll` 或漏导入某 Content 模块 | 补定义 / 补齐导入 |
| 文字变乱码 | 节点文字含中文 | `.bas` 是单字节文件，改英文/拼音；重存 UTF-8 也不可靠 |
| 图溢出幻灯片 | `PX_TO_IN`/`OFFSET` 算错 | 重算：scale 用 min，offset 用居中公式 |
| 箭头悬空 | 连线终点没贴形状边界 | 回图把 `AddPath` 末点移到形状边 |
| 单个模块卡顿 | 节点/连线过密没拆 | 把 `DrawContentN` 拆到 `modFlow_Content2.bas` |
| 颜色偏了 | 取色取成渐变中间值 | 图里用扁平色；渐变请取主色 |
| 生成的 `.pptx` 是空白的 | COM 未装上 PowerPoint，或 `BuildFlowchart` 抛错被吞 | 看脚本日志；改走回放路径，并回查 `.bas` 是否漏几何常量 |
| `review_render.py` FAIL | 字体过大溢出 / 悬空连线 / 穿模 / 出界 | 按输出清单逐条修 `.bas`，重跑 `build_ppt.py` 后再审 |
| 重建 `.pptx` 报 `PermissionError` | 上一轮 pptx 被预览面板/Office 占用，删除也失败 | `build_ppt.py <目录> --out <目录>\flowchart_v2.pptx` 换名输出；`review_render.py` 会自动选 mtime 最新的 pptx（`~$` 锁文件已过滤） |
| 字母堆叠竖排标签报"文本总高超出节点" | 多行按 1.25 行距累计，细高框装不下 | 字号 8pt→6.5pt（等效视觉密度）+ 透明框扩高保持中心；见 rubric 评审二 |
| 渲染图风格与原图"不像" | 配色/字号/线宽凭目测没程序化取值 | 回评审一：`extract_colors.py` + `compare_text.py` 重新测量 |
| 公式在 PPT 里打不开/显示线性文本 | latex2mathml / mathml2omml 未安装，或 fallback 匹配失败 | `pip install latex2mathml mathml2omml`；确认 AddFormula 的 fallback 文本与最终形状文本一致；COM 路径注入失败时保留线性版可用 |
| 公式溢出所在白框 | fontPt 过大（mathtext/PPT 宽度按 em 估算） | 缩 fontPt（公式占位宽 ≈ pt×1.83px/字符 × 字符数），或加宽 AddFormula 的 w |

## 资源

- `assets/modFlow_Engine.bas` — 固定引擎（坐标映射 / AddNode / AddPath / 文字 / 幂等 / 入口）
- `references/example_content.bas` — 可运行的内容模块示例
- `references/vba-module-map.md` — 引擎 API 全表 + 内容模块契约
- `references/review_rubric.md` — **两轮评审标准**（风格一致性 + 渲染几何检查）
- `references/review_examples/origin.jpg` — 风格基准原图（正样本）
- `references/review_examples/failed.png` — 失败案例渲染图（负样本，附八类典型失败）
- `scripts/build_ppt.py` — **最后一步**：`.bas` → 真实 `.pptx`（COM 优先，回放兜底）
- `scripts/lint_vba.py` — 交付前的结构自检
- `scripts/review_render.py` — **评审二程序化检查**：文本溢出 / 悬空连线 / 穿模 / 重叠 / 出界
- `scripts/render_preview.py` — 回放 `.bas` 渲染 `preview_compare.png`（与源图并排）
- `scripts/extract_colors.py` / `extract_geometry.py` / `compare_text.py` — 评审一程序化取色 / 取框 / 反推字号
