Attribute VB_Name = "modFlow_Engine"
' ============================================================================
' modFlow_Engine  -  fixed engine shipped by sci-flowchart-vba (do NOT edit)
' ============================================================================
' The skill ships this module; the model only writes modFlow_Content.bas.
' It owns: px->inch mapping, shape/path creation, text, tagging, idempotent
' cleanup and the entry points.
'
' modFlow_Content.bas MUST declare these Public Const (referenced here):
'   CANVAS_W_PX, CANVAS_H_PX   layout canvas size (px)
'   PX_TO_IN                   px -> inch scale factor
'   OFFSET_X_IN, OFFSET_Y_IN   top-left offset on the slide (inch)
'   SLIDE_W_IN, SLIDE_H_IN     chosen slide size (inch)
'   CUSTOM_SIZE                True when SLIDE_* differs from 13.333x7.5
'   FONT_NAME                  label font name (string)
' and MUST define: Public Sub DrawAll(sld As Slide)
' BuildFlowchart calls DrawAll, which calls every DrawContentN.
' ============================================================================
Option Explicit
Option Base 0

Public Const TAG_NAME As String = "SciFlow"

' mso constants as literals: avoids depending on library load order
Public Const ARROWHEAD_TRIANGLE As Long = 2
Public Const LINE_SOLID As Long = 1
Public Const LINE_DASH As Long = 4
Public Const LINE_DOT As Long = 2
Public Const LINE_DASHDOT As Long = 5

' ---- coordinate mapping (constants come from the content module) -----------
Public Function PX2X(ByVal px As Single) As Single
    PX2X = OFFSET_X_IN + px * PX_TO_IN
End Function

Public Function PX2Y(ByVal py As Single) As Single
    PX2Y = OFFSET_Y_IN + py * PX_TO_IN
End Function

Public Function PX2L(ByVal pl As Single) As Single
    PX2L = pl * PX_TO_IN
End Function

' ---- shape kind lookup table -----------------------------------------------
Public Function ShapeTypeFor(ByVal kind As String) As Long
    Select Case LCase$(kind)
        Case "rect", "rectangle", "box": ShapeTypeFor = 1        ' msoShapeRectangle
        Case "round_rect", "rounded", "rounded_rect": ShapeTypeFor = 5
        Case "oval", "ellipse", "circle": ShapeTypeFor = 9
        Case "diamond", "decision": ShapeTypeFor = 4
        Case "parallelogram": ShapeTypeFor = 2
        Case "trapezoid": ShapeTypeFor = 3
        Case "pentagon": ShapeTypeFor = 51
        Case "hexagon": ShapeTypeFor = 10
        Case "stadium", "pill", "terminator": ShapeTypeFor = 5   ' large corner radius
        Case "cylinder": ShapeTypeFor = 21
        Case "document": ShapeTypeFor = 61
        Case "right_arrow", "arrow_right": ShapeTypeFor = 33
        Case "left_arrow", "arrow_left": ShapeTypeFor = 34
        Case "up_arrow", "arrow_up": ShapeTypeFor = 35
        Case "down_arrow", "arrow_down": ShapeTypeFor = 36
        Case "chevron": ShapeTypeFor = 52
        Case Else: ShapeTypeFor = 1
    End Select
End Function

' ---- create a node ---------------------------------------------------------
Public Sub AddNode(ByVal sld As Slide, ByVal id As String, ByVal kind As String, _
                   ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, _
                   ByVal fillC As Long, ByVal lineC As Long, ByVal lineW As Single, _
                   ByVal dash As Long, _
                   Optional ByVal text As String = "", Optional ByVal fontPt As Single = 11, _
                   Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, _
                   Optional ByVal fontC As Long = -1, Optional ByVal cornerRatio As Single = -1, _
                   Optional ByVal fontName As String = "", Optional ByVal adj2 As Single = -1)
    Dim shp As Shape
    Dim st As Long
    st = ShapeTypeFor(kind)
    Set shp = sld.Shapes.AddShape(st, PX2X(x), PX2Y(y), PX2L(w), PX2L(h))
    With shp
        If fillC < 0 Then
            .Fill.Visible = msoFalse
        Else
            .Fill.Visible = msoTrue
            .Fill.Solid
            .Fill.ForeColor.RGB = fillC
        End If
        If lineC < 0 Then
            .Line.Visible = msoFalse
        Else
            .Line.Visible = msoTrue
            .Line.ForeColor.RGB = lineC
            .Line.Weight = lineW
            .Line.DashStyle = dash
        End If
    End With
    If LCase$(kind) = "stadium" Or LCase$(kind) = "pill" Or LCase$(kind) = "terminator" Then
        cornerRatio = 0.5
    End If
    ' cornerRatio -> Adjustments(1): corner radius ratio, or arrow head height
    If cornerRatio >= 0 Then
        On Error Resume Next
        shp.Adjustments.Item(1) = cornerRatio
        On Error GoTo 0
    End If
    ' adj2 -> Adjustments(2): block arrow head length ratio
    If adj2 >= 0 Then
        On Error Resume Next
        shp.Adjustments.Item(2) = adj2
        On Error GoTo 0
    End If
    If fontName = "" Then fontName = FONT_NAME
    If fontC < 0 Then fontC = RGB(31, 31, 31)
    SetLabel shp, text, fontPt, bold, italic, fontC, fontName
    TagShape shp, TAG_NAME, id
End Sub

' ---- create a connector (parses "x1,y1;x2,y2;..." in px) -------------------
Public Sub AddPath(ByVal sld As Slide, ByVal id As String, ByVal pts As String, _
                   ByVal lineC As Long, ByVal lineW As Single, ByVal dash As Long, _
                   ByVal arrowEnd As Boolean)
    Dim parts() As String
    Dim xy() As String
    Dim xs() As Single
    Dim ys() As Single
    Dim i As Long, n As Long
    If Len(pts) = 0 Then Exit Sub
    parts = Split(pts, ";")
    n = UBound(parts) + 1
    If n < 2 Then Exit Sub
    ReDim xs(0 To n - 1)
    ReDim ys(0 To n - 1)
    For i = 0 To n - 1
        xy = Split(parts(i), ",")
        If UBound(xy) >= 1 Then
            xs(i) = CSng(Trim$(xy(0)))
            ys(i) = CSng(Trim$(xy(1)))
        End If
    Next i
    DrawPolyline sld, xs, ys, n, lineC, lineW, dash, arrowEnd, id
End Sub

' ---- low level polyline (called by AddPath) --------------------------------
Public Sub DrawPolyline(ByVal sld As Slide, ptsX() As Single, ptsY() As Single, _
                        ByVal n As Long, ByVal lineC As Long, ByVal lineW As Single, _
                        ByVal dash As Long, ByVal arrowEnd As Boolean, ByVal id As String)
    Dim i As Long
    Dim seg As Shape
    If n < 2 Then Exit Sub
    For i = 0 To n - 2
        Set seg = sld.Shapes.AddLine(PX2X(ptsX(i)), PX2Y(ptsY(i)), _
                                     PX2X(ptsX(i + 1)), PX2Y(ptsY(i + 1)))
        With seg.Line
            .ForeColor.RGB = lineC
            .Weight = lineW
            .DashStyle = dash
            .BeginArrowheadStyle = msoArrowheadNone
            If arrowEnd And i = n - 2 Then
                .EndArrowheadStyle = msoArrowheadTriangle
                .EndArrowheadLength = msoArrowheadMedium
                .EndArrowheadWidth = msoArrowheadMedium
            Else
                .EndArrowheadStyle = msoArrowheadNone
            End If
        End With
        TagShape seg, TAG_NAME, id
    Next i
End Sub

' ---- write the label -------------------------------------------------------
Public Sub SetLabel(ByVal shp As Shape, ByVal txt As String, _
                    ByVal sizePt As Single, ByVal bold As Boolean, _
                    ByVal italic As Boolean, ByVal fontC As Long, _
                    ByVal fontName As String)
    If Len(txt) = 0 Then Exit Sub
    With shp.TextFrame
        .WordWrap = msoTrue
        .AutoSize = msoAutoSizeNone
        .MarginLeft = 1
        .MarginRight = 1
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = txt
        .TextRange.Font.Name = fontName
        .TextRange.Font.Size = sizePt
        .TextRange.Font.Bold = bold
        .TextRange.Font.Italic = italic
        .TextRange.Font.Color.RGB = fontC
        .TextRange.ParagraphFormat.Alignment = ppAlignCenter
        .TextRange.ParagraphFormat.SpaceBefore = 0
        .TextRange.ParagraphFormat.SpaceAfter = 0
        .TextRange.ParagraphFormat.LineRuleWithin = msoTrue
        .TextRange.ParagraphFormat.SpaceWithin = 0.95
    End With
    On Error Resume Next
    shp.TextFrame2.AutoSize = msoAutoSizeTextToFitShape
    shp.TextFrame2.AutoSize = msoAutoSizeNone
    On Error GoTo 0
End Sub

' ---- rich-text helpers (content modules may call these; the replay and
'      build_ppt.py parse the SAME calls so both paths render alike) -----------

' Find a tagged shape by id (first match)
Public Function ShapeById(ByVal sld As Slide, ByVal id As String) As Shape
    Dim shp As Shape
    For Each shp In sld.Shapes
        If shp.Tags(TAG_NAME) = id Then
            Set ShapeById = shp
            Exit Function
        End If
    Next shp
End Function

' Paragraph alignment for a tagged shape: align = "left" | "center" | "right";
' marginPx is the extra left margin expressed in canvas px.
Public Sub SetPara(ByVal sld As Slide, ByVal id As String, ByVal align As String, _
                   ByVal marginPx As Single)
    Dim shp As Shape
    Set shp = ShapeById(sld, id)
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    Select Case LCase$(align)
        Case "left": shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        Case "right": shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
        Case Else: shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End Select
    shp.TextFrame2.MarginLeft = PX2L(marginPx) * 72
    On Error GoTo 0
End Sub

' Color + bold the FIRST nChars characters of a tagged shape's label
' (e.g. the red "Step 1:" prefix inside a banner).
Public Sub SetPartColor(ByVal sld As Slide, ByVal id As String, ByVal nChars As Long, _
                        ByVal colorC As Long)
    Dim shp As Shape
    Set shp = ShapeById(sld, id)
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    With shp.TextFrame2.TextRange.Characters(1, nChars).Font
        .Fill.ForeColor.RGB = colorC
        .Bold = msoTrue
    End With
    On Error GoTo 0
End Sub

' ---- tags / idempotency ----------------------------------------------------
Public Sub ClearPrevious(ByVal sld As Slide, ByVal tagName As String)
    Dim i As Long
    Dim shp As Shape
    For i = sld.Shapes.Count To 1 Step -1
        Set shp = sld.Shapes(i)
        If HasTag(shp, tagName) Then shp.Delete
    Next i
End Sub

Public Sub TagShape(ByVal shp As Shape, ByVal tagName As String, ByVal tagValue As String)
    On Error Resume Next
    shp.Tags.Add tagName, tagValue
    On Error GoTo 0
End Sub

Public Function HasTag(ByVal shp As Shape, ByVal tagName As String) As Boolean
    Dim t As String
    HasTag = False
    On Error Resume Next
    t = shp.Tags(tagName)
    On Error GoTo 0
    HasTag = (Len(t) > 0)
End Function

' ---- entry points ----------------------------------------------------------
Public Sub BuildFlowchart()
    Dim sld As Slide
    Dim tgt As Presentation
    On Error GoTo Fail
    On Error Resume Next
    Set sld = Application.ActiveWindow.View.Slide
    On Error GoTo Fail
    If sld Is Nothing Then
        Set tgt = Application.ActivePresentation
        Set sld = tgt.Slides.Add(1, ppLayoutBlank)
    End If
    Application.DisplayAlerts = ppAlertsNone
    Application.ScreenUpdating = False
    If CUSTOM_SIZE Then
        Application.ActivePresentation.PageSetup.SlideWidth = SLIDE_W_IN * 72
        Application.ActivePresentation.PageSetup.SlideHeight = SLIDE_H_IN * 72
    End If
    ClearPrevious sld, TAG_NAME
    DrawAll sld
    Application.ScreenUpdating = True
    Application.DisplayAlerts = ppAlertsAll
    MsgBox "Flowchart built: " & sld.Shapes.Count & " shapes.", vbInformation, "sci-flowchart-vba"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = ppAlertsAll
    MsgBox "Build failed (" & Err.Number & "): " & Err.Description, vbExclamation, "sci-flowchart-vba"
End Sub

Public Sub RemoveFlowchart()
    Dim sld As Slide
    On Error Resume Next
    Set sld = Application.ActiveWindow.View.Slide
    On Error GoTo 0
    If sld Is Nothing Then Exit Sub
    ClearPrevious sld, TAG_NAME
End Sub
