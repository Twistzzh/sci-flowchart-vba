# 模块职责与契约（vision-direct 版）

新版管线**完全去掉 SVG 矢量化与三脚本管线**：多模态大模型直接看图，
产出 VBA 内容模块；坐标映射、形状创建、连线、文字、幂等清理、入口过程
都由固定引擎 `assets/modFlow_Engine.bas` 保证。大模型只写"内容"。

## 交付物结构

| 文件 | 来源 | 职责 |
|---|---|---|
| `modFlow_Engine.bas` | skill 原样提供（**勿改**） | 引擎：坐标映射、AddNode/AddPath、文字、打标签、幂等清理、BuildFlowchart/RemoveFlowchart |
| `modFlow_Content.bas` | **大模型从图/论文生成** | 几何常量 + 调色板 + `DrawAll` + `DrawContentN`（节点与连线） |
| `modFlow_Content2.bas`（可选） | 大模型生成 | 图太大时把额外 `DrawContentN` 拆到这里 |
| `flowchart.pptx` | `scripts/build_ppt.py` 自动生成 | **最终交付物**：形状已画好、可直接打开编辑 |

导入顺序任意；模块之间靠过程名调用。

## 最后一步：`.bas` → `.pptx`

`scripts/build_ppt.py <输出目录>` 负责把上面几个 `.bas` 变成真正能打开的 PPT：

1. 按固定顺序装配模块：`modFlow_Engine` 先、`modFlow_Content*` 后
   （内容模块要先于引擎编译会找不到 `DrawAll` 的引用关系，反之则常量未定义）。
2. 新建空白 16:9 演示文稿；若内容模块 `CUSTOM_SIZE=True`，按
   `SLIDE_W_IN / SLIDE_H_IN` 调整页面尺寸。
3. **真的执行 `BuildFlowchart`**：COM 路径下由 PowerPoint 引擎亲自绘制，
   与用户手动 Run 的结果一致；随后 `SaveAs` 出 `.pptx`（并留 `.pptm` 保宏）。
4. 无 PowerPoint 时降级为内置回放器：解析 `references/vba-module-map.md`
   里描述的同一份 `.bas` 语义（AddNode/AddPath 实参），用 `python-pptx`
   生成等价原生形状。两条路径的形状/配色/坐标语义一致。

**产出校验**：脚本结束前必须确认幻灯片上 `Shapes.Count > 0` 且与
`AddNode + AddPath` 的调用条数匹配，否则视为失败（说明有内容没画出来）。

## 引擎 API（内容模块调用这些）

| 过程 | 签名 | 作用 |
|---|---|---|
| `PX2X` / `PX2Y` / `PX2L` | `Function(px As Single) As Single` | px → 幻灯片英寸（坐标 / 长度） |
| `ShapeTypeFor` | `Function(kind As String) As Long` | 形状名 → mso 类型（见下） |
| `AddNode` | `Sub(sld, id, kind, x, y, w, h, fillC, lineC, lineW, dash, [text], [fontPt], [bold], [italic], [fontC], [cornerRatio], [fontName], [adj2])` | 建一个节点并填字、打标签 |
| `AddPath` | `Sub(sld, id, pts$, lineC, lineW, dash, arrowEnd)` | 按点串 `"x1,y1;x2,y2;..."` 画折线连线（末段带箭头） |
| `DrawPolyline` | 底层，被 AddPath 调用 | 逐段 AddLine |
| `SetPara` | `Sub(sld, id, align$, marginPx!)` | 事后排版：`align` = `"left"/"center"/"right"`，`marginPx` 为文字左内边距（px）。**必须在全部 AddNode 之后调用**（可跨模块引用别处定义的节点） |
| `SetPartColor` | `Sub(sld, id, nChars&, colorC&)` | 把节点文字前 `nChars` 个字符单独标色 + 加粗（如 `"Step 1:"` 标红）。同样必须在 AddNode 之后调用 |
| `SetLabel` / `StyleShape` / `TagShape` / `HasTag` / `ClearPrevious` | 工具 | 文字/样式/标签/幂等 |
| `BuildFlowchart` / `RemoveFlowchart` | 入口 | 运行/清空 |

`cornerRatio` → `Adjustments(1)`：圆角矩形的圆角比例，或块箭头的箭头相对高度。
`adj2` → `Adjustments(2)`：仅块箭头用，箭头相对长度。只给块箭头参数时，中间的
`text/fontPt/bold/italic/fontC/fontName` 槽位必须补上 `""`、`11`、`False`、`False`、`-1`、`""`，
**不能留空实参**（VBA 不接受 `f(a, , b)` 之外的省略写法在此处的可读性）。

### 形状 kind 字符串 → mso 类型

`rect`(1) · `round_rect`(5) · `oval`(9) · `diamond`(4) · `parallelogram`(2) ·
`trapezoid`(3) · `pentagon`(51) · `hexagon`(10) · `stadium`/`pill`/`terminator`(5, 自动大圆角) ·
`cylinder`(21) · `document`(61) · `right_arrow`(33) · `left_arrow`(34) · `up_arrow`(35) ·
`down_arrow`(36) · `chevron`(52)。未识别 → `rect`。

### 颜色 / 线型字面量（引擎已定义）

`LINE_SOLID=1` · `LINE_DASH=4` · `LINE_DOT=2` · `LINE_DASHDOT=5`。
颜色一律 `RGB(r,g,b)`。填充/描边传 `-1` 表示"无"。

## 内容模块必须满足的契约

1. **声明几何常量**（引擎引用，缺一则编译失败）：
   `CANVAS_W_PX, CANVAS_H_PX, PX_TO_IN, OFFSET_X_IN, OFFSET_Y_IN,
    SLIDE_W_IN, SLIDE_H_IN, CUSTOM_SIZE, FONT_NAME`。
   `FONT_NAME` 取与源图字形一致的字体族：西文衬线 `Times New Roman` /
   西文无衬线 `Arial`；中文黑体/无衬线 `Microsoft YaHei`、中文宋体/衬线
   `SimSun`、粗黑 `SimHei`。预览与评审脚本按它自动选 CJK 字体文件；
   个别节点需要不同字体时用 `AddNode` 的 `fontName` 可选参数覆盖。
2. **声明调色板** `Public Const`：`INK / BG / ACCENT / ACCENT2 / ALT / LINE` 等。
3. **定义唯一入口**：`Public Sub DrawAll(sld As Slide)`，内部调用各 `DrawContentN`。
4. 复杂图（节点 > 18 或单模块 > 300 行）把 `DrawContentN` 拆到 `modFlow_Content2.bas` 等，
   `DrawAll` 仍要依次调用它们（跨模块 Public Sub 全局可见）。
5. **整个文件 UTF-8 编码 + CRLF 换行**：节点文字、注释都可以直接用中文，
   源图是中文流程图时**直接保留原中文**（不要翻译成英文/拼音）。UTF-8 才能让中文
   正常落地；行尾必须是 CRLF。需要两行文字写 `"A" & vbLf & "B"`，比自动折行更可控。
6. 坐标、颜色、形状、文字**全部来自对源图的识别**，禁止凭空猜测。
7. **装饰调用统一放在 `Decorate` 私有 Sub 里**（`DrawAll` 末尾调用）：所有
   `SetPara` / `SetPartColor` 必须写在全部 AddNode 之后执行的位置——回放器
   与 lint 都是两遍解析（先收集节点，再应用装饰），写错顺序会匹配不到节点。
   覆盖式标题（文字叠在列表框上方）的避让写法：列表正文前导空行
   `"" & vbLf & "-Item..."`，标题用无边框无填充 rect 叠加，两条渲染路径都成立。
8. 交付前用 `scripts/lint_vba.py <输出目录>` 过一遍（常量/入口/空实参/编码与换行/点串格式）。
9. 交付前用 `scripts/build_ppt.py <输出目录>` 生成 `.pptx`；交付以 `.pptx` 为主、
   `.bas` 为辅（`.bas` 是"可重建的源码"，`.pptx` 是用户实际要用的文件）。

## 修改生成的 VBA

- 改配色 → 改 `modFlow_Content.bas` 里的 `Public Const`。
- 改位置/尺寸 → 改 `AddNode` / `AddPath` 的实参（像素语义）。
- 改字号 → 改 `AddNode` 的 `fontPt` 实参（已是 pt）。
- 重跑安全：所有形状带 `SciFlow` 标签，重复 Run 先清后建，不叠图。
- 别删 `modFlow_Engine.bas`：内容模块全部依赖它。
- 导入报"过程名未定义"通常是漏导入了 `modFlow_Engine.bas` 或某个 `Content` 模块。

完整可运行示例见 `references/example_content.bas`。
