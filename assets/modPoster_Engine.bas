Attribute VB_Name = "modPoster_Engine"
' ============================================================================
' modPoster_Engine  -  fixed engine shipped by poster-vba (do NOT edit)
' ============================================================================
' Poster/board variant of the sci-flowchart engine. The skill ships this
' module; the model only writes modPoster_Content.bas.
' It owns: px->inch mapping, shape/path/picture/background/bars creation,
' text (latin + East-Asian font), tagging, idempotent cleanup, entry points.
'
' modPoster_Content.bas MUST declare these Public Const (referenced here):
'   CANVAS_W_PX, CANVAS_H_PX   layout canvas size (px)
'   PX_TO_IN                   px -> inch scale factor
'   OFFSET_X_IN, OFFSET_Y_IN   top-left offset on the slide (inch)
'   SLIDE_W_IN, SLIDE_H_IN     chosen slide size (inch)
'   CUSTOM_SIZE                True when SLIDE_* differs from 13.333x7.5
'   FONT_NAME                  latin font family (string)
'   FONT_NAME_CN               East-Asian font family (string)
'   ASSET_DIR                  absolute folder holding icon images,
'                              with trailing backslash, e.g. "C:\out\assets\"
' and MUST define: Public Sub DrawAll(sld As Slide)
' BuildPoster calls DrawAll, which calls every DrawContentN.
'
' DRAW ORDER == CALL ORDER in DrawAll (background first, icons above panels).
' The replay path in build_poster.py preserves the same order.
' ============================================================================
Option Explicit
Option Base 0

Public Const TAG_NAME As String = "SciPoster"

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
' Values verified against the MSO_AUTO_SHAPE_TYPE enumeration.
Public Function ShapeTypeFor(ByVal kind As String) As Long
    Select Case LCase$(kind)
        ' --- flowchart staples -------------------------------------------------
        Case "rect", "rectangle", "box": ShapeTypeFor = 1        ' msoShapeRectangle
        Case "round_rect", "rounded", "rounded_rect": ShapeTypeFor = 5
        Case "oval", "ellipse", "circle": ShapeTypeFor = 9
        Case "diamond", "decision": ShapeTypeFor = 4
        Case "parallelogram": ShapeTypeFor = 2
        Case "trapezoid": ShapeTypeFor = 3
        Case "pentagon": ShapeTypeFor = 51                       ' home-plate pentagon
        Case "pentagon_reg": ShapeTypeFor = 12                   ' regular pentagon
        Case "hexagon": ShapeTypeFor = 10
        Case "stadium", "pill", "terminator": ShapeTypeFor = 5   ' large corner radius
        Case "can", "cylinder": ShapeTypeFor = 13                ' msoShapeCan
        Case "document": ShapeTypeFor = 61
        Case "right_arrow", "arrow_right": ShapeTypeFor = 33
        Case "left_arrow", "arrow_left": ShapeTypeFor = 34
        Case "up_arrow", "arrow_up": ShapeTypeFor = 35
        Case "down_arrow", "arrow_down": ShapeTypeFor = 36
        Case "chevron": ShapeTypeFor = 52
        Case "arc": ShapeTypeFor = 25
        ' --- playful / cartoon extras -----------------------------------------
        Case "star4": ShapeTypeFor = 91
        Case "star5": ShapeTypeFor = 92
        Case "star8": ShapeTypeFor = 93
        Case "star6": ShapeTypeFor = 147
        Case "heart": ShapeTypeFor = 21
        Case "cloud": ShapeTypeFor = 179
        Case "sun": ShapeTypeFor = 23
        Case "moon": ShapeTypeFor = 24
        Case "lightning", "bolt": ShapeTypeFor = 22
        Case "smiley": ShapeTypeFor = 17
        Case "donut": ShapeTypeFor = 18
        Case "wave": ShapeTypeFor = 103
        Case "double_wave": ShapeTypeFor = 104
        Case "callout", "round_callout": ShapeTypeFor = 106      ' rounded rectangular callout
        Case "bubble", "oval_callout": ShapeTypeFor = 107        ' oval speech bubble
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
    Dim latinName As String
    Dim eaName As String
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
    ' cornerRatio -> Adjustments(1): corner radius ratio, star inner radius, etc.
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
    If fontName = "" Then
        latinName = FONT_NAME
        eaName = FONT_NAME_CN
    Else
        latinName = fontName
        eaName = fontName
    End If
    If fontC < 0 Then fontC = RGB(31, 31, 31)
    SetLabel shp, text, fontPt, bold, italic, fontC, latinName, eaName
    TagShape shp, TAG_NAME, id
End Sub

' ---- insert a picture (icon / decorative cut-out) --------------------------
' fileName is resolved against: 1) as given (absolute path), 2) ASSET_DIR.
' w/h in canvas px; pass w<=0 And h<=0 for the image's natural size
' (prefer explicit sizes: it keeps the two build paths identical).
' A missing file degrades to a dashed placeholder box - the build never
' hard-fails on a lost asset.
Public Sub AddPicture(ByVal sld As Slide, ByVal id As String, ByVal fileName As String, _
                      ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Dim p As String
    Dim shp As Shape
    p = ResolveAsset(fileName)
    On Error GoTo Missing
    If w <= 0 And h <= 0 Then
        Set shp = sld.Shapes.AddPicture(p, msoFalse, msoTrue, PX2X(x), PX2Y(y))
    Else
        Set shp = sld.Shapes.AddPicture(p, msoFalse, msoTrue, PX2X(x), PX2Y(y), PX2L(w), PX2L(h))
    End If
    TagShape shp, TAG_NAME, id
    Exit Sub
Missing:
    On Error GoTo 0
    Dim pw As Single
    Dim ph As Single
    pw = w: If pw <= 0 Then pw = 120
    ph = h: If ph <= 0 Then ph = 90
    AddNode sld, id & "_ph", "rect", x, y, pw, ph, RGB(238, 238, 238), RGB(170, 170, 170), _
                1, LINE_DASH, "IMG " & fileName, 9, False, False, RGB(120, 120, 120)
End Sub

Private Function ResolveAsset(ByVal fileName As String) As String
    Dim p As String
    p = fileName
    If Len(Dir(p)) > 0 Then
        ResolveAsset = p
        Exit Function
    End If
    p = ASSET_DIR & fileName
    If Len(Dir(p)) > 0 Then
        ResolveAsset = p
        Exit Function
    End If
    ResolveAsset = fileName   ' keep original: AddPicture raises -> placeholder
End Function

' ---- full-bleed background -------------------------------------------------
' c2 < 0 -> solid c1; otherwise a two-color gradient c1 (top) -> c2 (bottom).
' Call AddBG FIRST inside DrawAll so every other shape sits above it.
Public Sub AddBG(ByVal sld As Slide, ByVal id As String, ByVal c1 As Long, ByVal c2 As Long)
    Dim shp As Shape
    Set shp = sld.Shapes.AddShape(1, 0, 0, SLIDE_W_IN * 72, SLIDE_H_IN * 72)
    With shp
        .Line.Visible = msoFalse
        .Fill.Solid
        .Fill.ForeColor.RGB = c1
        If c2 >= 0 Then
            .Fill.BackColor.RGB = c2
            Call .Fill.TwoColorGradient(msoGradientHorizontal, 1)
        End If
    End With
    TagShape shp, TAG_NAME, id
End Sub

' ---- mini column chart drawn with native shapes ----------------------------
' A lightweight, dependency-free chart for posters ("just a simple one").
' cats/vals are comma strings of equal length. Geometry (canvas px):
'   title band   : y .. y + 0.16*h            (only when title <> "")
'   cat band     : y + 0.86*h .. y + h        (height 0.14*h)
'   baseline     : y + 0.86*h
'   slot = w/n; bar width = 0.6*slot; bar i left = x + slot*i + 0.2*slot
'   bar i height = vals(i)/vmax * plotH * 0.90 (vmax = max(vals), min 1)
'   value label  : above the bar, band height 0.08*h, capped inside plot
' build_poster.py replay mirrors this spec exactly - do not change one side.
Public Sub AddBars(ByVal sld As Slide, ByVal id As String, _
                   ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, _
                   ByVal title As String, ByVal cats As String, ByVal vals As String, _
                   ByVal barC As Long, ByVal lineC As Long, ByVal txtC As Long, _
                   Optional ByVal labelPt As Single = 9, Optional ByVal titlePt As Single = 11, _
                   Optional ByVal fontName As String = "")
    Dim vc() As String
    Dim cc() As String
    Dim v() As Single
    Dim i As Long
    Dim n As Long
    Dim vmax As Single
    Dim titleH As Single
    Dim plotTop As Single
    Dim catH As Single
    Dim ybase As Single
    Dim plotH As Single
    Dim slot As Single
    Dim bx As Single
    Dim bh As Single
    Dim by As Single
    Dim ly As Single
    vc = Split(vals, ",")
    n = UBound(vc) + 1
    If n < 1 Then Exit Sub
    cc = Split(cats & ",", ",")
    ReDim v(0 To n - 1)
    vmax = 0
    For i = 0 To n - 1
        v(i) = CSng(Trim$(vc(i)))
        If v(i) > vmax Then vmax = v(i)
    Next i
    If vmax <= 0 Then vmax = 1
    titleH = 0: If Len(title) > 0 Then titleH = 0.16 * h
    plotTop = y + titleH
    catH = 0.14 * h
    ybase = y + h - catH
    plotH = ybase - plotTop
    slot = w / n
    If Len(title) > 0 Then
        AddNode sld, id & "_t", "rect", x, y, w, titleH, -1, -1, 1, LINE_SOLID, _
                    title, titlePt, True, False, txtC
    End If
    For i = 0 To n - 1
        bx = x + slot * i + 0.2 * slot
        bh = (v(i) / vmax) * plotH * 0.9
        If bh > plotH Then bh = plotH
        by = ybase - bh
        AddNode sld, id & "_b" & (i + 1), "rect", bx, by, 0.6 * slot, bh, barC, -1, 1, LINE_SOLID, "", 8
        ly = by - 0.08 * h
        If ly < plotTop Then ly = plotTop
        AddNode sld, id & "_v" & (i + 1), "rect", bx - 0.1 * slot, ly, slot, 0.08 * h, _
                    -1, -1, 1, LINE_SOLID, FormatNum(v(i)), labelPt, False, False, txtC
        If i < UBound(cc) + 1 Then
            If Len(Trim$(cc(i))) > 0 Then
                AddNode sld, id & "_c" & (i + 1), "rect", x + slot * i, ybase, slot, catH, _
                            -1, -1, 1, LINE_SOLID, Trim$(cc(i)), labelPt, False, False, txtC
            End If
        End If
    Next i
    AddPath sld, id & "_axis", CStr(x) & "," & CStr(ybase) & ";" & CStr(x + w) & "," & CStr(ybase), _
                lineC, 1, LINE_SOLID, False
End Sub

Private Function FormatNum(ByVal v As Single) As String
    Dim s As String
    s = Format$(v, "0.##")
    If Len(s) = 0 Then s = "0"
    FormatNum = s
End Function

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

' ---- write the label (latin font + East-Asian font) ------------------------
Public Sub SetLabel(ByVal shp As Shape, ByVal txt As String, _
                    ByVal sizePt As Single, ByVal bold As Boolean, _
                    ByVal italic As Boolean, ByVal fontC As Long, _
                    ByVal fontName As String, ByVal eaName As String)
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
        On Error Resume Next
        .TextRange.Font.NameFarEast = eaName   ' Chinese glyphs use this family
        On Error GoTo 0
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

' ---- rich-text helpers ------------------------------------------------------

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
' (e.g. the red keyword prefix inside a banner).
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

' Move a tagged shape to the front (call AFTER DrawAll content, e.g. at the
' very end of DrawAll, when an icon must sit above later-drawn panels).
Public Sub BringToFront(ByVal sld As Slide, ByVal id As String)
    Dim shp As Shape
    Set shp = ShapeById(sld, id)
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    shp.ZOrder msoBringToFront
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
Public Sub BuildPoster()
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
    MsgBox "Poster built: " & sld.Shapes.Count & " shapes.", vbInformation, "poster-vba"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = ppAlertsAll
    MsgBox "Build failed (" & Err.Number & "): " & Err.Description, vbExclamation, "poster-vba"
End Sub

Public Sub RemovePoster()
    Dim sld As Slide
    On Error Resume Next
    Set sld = Application.ActiveWindow.View.Slide
    On Error GoTo 0
    If sld Is Nothing Then Exit Sub
    ClearPrevious sld, TAG_NAME
End Sub
