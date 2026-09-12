Attribute VB_Name = "modFlow_Content"
' ============================================================================
' modFlow_Content  -  示例内容模块（由大模型从图片/论文读取后编写）
' 本文件示范完整契约：几何常量 + 调色板 + DrawAll + DrawContentN。
' 真实生成时，坐标、颜色、文字、形状全都来自对源图的识别，不得凭空编造。
' 文字直接用源语言即可：源图是中文流程图时，节点文字就直接写中文（见下方 "数据收集"）。
' ============================================================================
Option Explicit
Option Base 0

' --- 几何（由源图像尺寸推算；下面为示例数值）------------------------------
Public Const CANVAS_W_PX As Single = 800
Public Const CANVAS_H_PX As Single = 500
Public Const PX_TO_IN As Single = 0.0137
Public Const OFFSET_X_IN As Single = 1.18
Public Const OFFSET_Y_IN As Single = 0.32
Public Const SLIDE_W_IN As Single = 13.333
Public Const SLIDE_H_IN As Single = 7.5
Public Const CUSTOM_SIZE As Boolean = False
Public Const FONT_NAME As String = "Arial"

' --- 调色板（RGB 字面量，来自对图取色）------------------------------------
Public Const INK As Long = RGB(31, 31, 31)
Public Const BG As Long = RGB(255, 255, 255)
Public Const ACCENT As Long = RGB(68, 114, 196)
Public Const ALT As Long = RGB(226, 237, 250)
Public Const LINE As Long = RGB(120, 120, 120)

' --- 入口：BuildFlowchart 会调用它 -----------------------------------------
Public Sub DrawAll(ByVal sld As Slide)
    DrawContent1 sld
End Sub

' --- 第 1 批内容 ------------------------------------------------------------
Public Sub DrawContent1(ByVal sld As Slide)
    ' 节点：AddNode sld, id, kind, x, y, w, h, fill, line, lineW, dash, _
    '                  text, fontPt, bold, italic, fontC, cornerRatio, fontName
    AddNode sld, "start", "stadium", 340, 20, 120, 50, ACCENT, -1, 1, LINE_SOLID, _
                "Start", 14, True, False, INK
    AddNode sld, "a", "round_rect", 320, 120, 160, 60, BG, LINE, 1, LINE_SOLID, _
                "数据收集", 12, False, False, INK, 0.12
    AddNode sld, "dec", "diamond", 350, 230, 100, 90, ALT, LINE, 1, LINE_SOLID, _
                "Good?", 12, False, False, INK
    AddNode sld, "b", "round_rect", 520, 245, 160, 60, BG, LINE, 1, LINE_SOLID, _
                "Train model", 12, False, False, INK, 0.12
    AddNode sld, "end", "stadium", 340, 400, 120, 50, ACCENT, -1, 1, LINE_SOLID, _
                "Finish", 14, True, False, INK

    ' 连线：AddPath sld, id, "x1,y1;x2,y2;...", line, lineW, dash, arrowEnd
    AddPath sld, "e1", "400,70;400,120", LINE, 1, LINE_SOLID, True
    AddPath sld, "e2", "400,180;400,230", LINE, 1, LINE_SOLID, True
    AddPath sld, "e3", "450,275;520,275", LINE, 1, LINE_SOLID, True
    AddPath sld, "e4", "400,320;400,400", LINE, 1, LINE_SOLID, True
End Sub
