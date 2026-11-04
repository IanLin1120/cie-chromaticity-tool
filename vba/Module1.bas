Attribute VB_Name = "Module1"
' CIE 1931 xy measurement report - recalculation / chart macros
' Source is ASCII-only on purpose (Chinese text is built with ChrW so it works with any code page).
Option Explicit

Private Const R0 As Long = 21          ' first data row (row 20 = header)
Private Const C_IMG As String = "J2"   ' image selector (data validation list)
Private Const C_BOX As String = "K2"   ' check cells (first char is a check mark)
Private Const C_PTS As String = "L2"
Private Const C_AVG As String = "M2"
Private Const FONT_NAME As String = "Microsoft JhengHei"

' current image transform (set by Overlay)
Private mVX0 As Double, mVX1 As Double, mVY0 As Double, mVY1 As Double
Private mFL As Double, mFT As Double, mFW As Double, mFH As Double
Private mIL As Double, mIT As Double, mIW As Double, mIH As Double
Private mDot As Double, mBoxW As Double, mBoxCol As Long
Private mNames() As String, mCnt As Long      ' names of the overlay shapes created for the current image

' ---------- text helpers (Traditional Chinese via ChrW) ----------
Private Function T_OOS() As String
    T_OOS = {{超規}}
End Function
Private Function T_DONE() As String
    T_DONE = {{已重新計算完成。量測筆數：}}
End Function
Private Function T_BAD() As String
    T_BAD = {{，超規筆數：}}
End Function
Private Function T_ERR() As String
    T_ERR = {{重新計算時發生錯誤：}}
End Function
Private Function T_NOSHEET() As String
    T_NOSHEET = {{找不到量測數據工作表（A20 應為 No.）}}
End Function
Private Function T_NOSPEC() As String
    T_NOSPEC = {{未設定規格}}
End Function
Private Function T_TITLE() As String
    T_TITLE = {{CIE 色度量測}}
End Function
Private Function CHK() As String
    CHK = ChrW(&H2611)
End Function
Private Function UNCHK() As String
    UNCHK = ChrW(&H2610)
End Function

' ---------- utilities ----------
Private Function IsNum(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Trim$(v) = "" Then Exit Function
    End If
    IsNum = IsNumeric(v)
End Function

Private Function LastRowIn(ByVal ws As Worksheet, ByVal c1 As Long, ByVal c2 As Long) As Long
    Dim c As Long, r As Long
    LastRowIn = R0 - 1
    For c = c1 To c2
        r = ws.Cells(ws.Rows.Count, c).End(-4162).Row   ' xlUp
        If r > LastRowIn Then LastRowIn = r
    Next c
End Function

Private Function ParamSheet() As Worksheet
    Dim s As Worksheet
    For Each s In ThisWorkbook.Worksheets
        If s.Name = "_cie_params" Then
            Set ParamSheet = s
            Exit Function
        End If
    Next s
End Function

' global settings: key in column L, value in column M of the parameter sheet
Private Function Gp(ByVal key As String, ByVal dflt As Double) As Double
    Dim ps As Worksheet, r As Long
    Gp = dflt
    Set ps = ParamSheet()
    If ps Is Nothing Then Exit Function
    For r = 1 To 60
        If CStr(ps.Cells(r, 12).Value) = key Then
            If IsNum(ps.Cells(r, 13).Value) Then Gp = CDbl(ps.Cells(r, 13).Value)
            Exit Function
        End If
    Next r
End Function

Private Function FindDataSheet() As Worksheet
    Dim s As Worksheet
    Set s = ActiveSheet
    If Not s Is Nothing Then
        If CStr(s.Range("A20").Value) = "No." Then
            Set FindDataSheet = s
            Exit Function
        End If
    End If
    For Each s In ThisWorkbook.Worksheets
        If CStr(s.Range("A20").Value) = "No." Then
            Set FindDataSheet = s
            Exit Function
        End If
    Next s
End Function

Private Sub MarkBad(ByVal rg As Range, ByVal withFill As Boolean)
    rg.Font.Color = RGB(255, 0, 0)
    rg.Font.Bold = True
    If withFill Then rg.Interior.Color = RGB(253, 226, 225)
End Sub

Private Sub AddU(ByRef u As Range, ByVal c As Range)
    If u Is Nothing Then
        Set u = c
    Else
        Set u = Application.Union(u, c)
    End If
End Sub

' fit columns A:G to their content (merged title rows are ignored by AutoFit), keep a little padding
Private Sub FitColumns(ByVal ws As Worksheet)
    Dim c As Long, w As Double
    ws.Range("A:G").Columns.AutoFit
    For c = 1 To 7
        w = ws.Columns(c).ColumnWidth + 1.5
        If w < 7 Then w = 7
        If w > 40 Then w = 40
        ws.Columns(c).ColumnWidth = w
    Next c
End Sub

Private Sub ClearMark(ByVal rg As Range)
    rg.Font.Color = RGB(0, 0, 0)
    rg.Font.Bold = False
    rg.Interior.Pattern = -4142   ' xlNone
End Sub

Public Function IsOn(ByVal c As Range) As Boolean
    IsOn = (Left$(CStr(c.Value), 1) = CHK())
End Function

Public Sub ToggleBox(ByVal c As Range)
    Dim t As String
    t = CStr(c.Value)
    If Len(t) = 0 Then t = UNCHK()
    If IsOn(c) Then
        c.Value = UNCHK() & Mid$(t, 2)
    Else
        c.Value = CHK() & Mid$(t, 2)
    End If
End Sub

' ---------- main: assigned to the button ----------
Public Sub RecalcAll()
    Dim ws As Worksheet
    Set ws = FindDataSheet()
    If ws Is Nothing Then
        MsgBox T_NOSHEET(), vbExclamation, T_TITLE()
        Exit Sub
    End If

    Dim calcMode As Long, evts As Boolean
    calcMode = Application.Calculation
    evts = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = -4135   ' xlCalculationManual
    On Error Resume Next
    ws.DisplayPageBreaks = False      ' page breaks slow down shape operations
    On Error GoTo Fail

    Dim r As Long, lastR As Long, lastAny As Long, n As Long, nBad As Long

    ' 1) remove rows that no longer have x and y (deleted / cleared measurements)
    Dim delRg As Range
    lastAny = LastRowIn(ws, 1, 7)
    For r = lastAny To R0 Step -1
        If Trim$(CStr(ws.Cells(r, 3).Value)) = "" And Trim$(CStr(ws.Cells(r, 4).Value)) = "" Then AddU delRg, ws.Rows(r)
    Next r
    If Not delRg Is Nothing Then delRg.EntireRow.Delete
    lastR = LastRowIn(ws, 3, 4)
    If lastR < R0 Then n = 0 Else n = lastR - R0 + 1

    ' 2) spec limits
    Dim sx0 As Variant, sx1 As Variant, sy0 As Variant, sy1 As Variant, specOK As Boolean
    sx0 = ws.Range("B6").Value: sx1 = ws.Range("C6").Value
    sy0 = ws.Range("B7").Value: sy1 = ws.Range("C7").Value
    specOK = IsNum(sx0) And IsNum(sx1) And IsNum(sy0) And IsNum(sy1)
    If specOK Then
        ws.Range("D6").Formula = "=AVERAGE(B6:C6)"
        ws.Range("D7").Formula = "=AVERAGE(B7:C7)"
    End If

    ' 3) numbering, judgement formulas and highlighting (whole block at once)
    If n > 0 Then
        Dim blk As Range, nums() As Variant, vals As Variant, i As Long
        Dim x As Variant, y As Variant, bx As Boolean, by As Boolean
        Dim badV As Range, badJ As Range, badG As Range
        Set blk = ws.Range(ws.Cells(R0, 1), ws.Cells(lastR, 7))
        ClearMark blk
        blk.Font.Name = FONT_NAME
        blk.Borders.LineStyle = 1                   ' xlContinuous
        blk.Borders.Color = RGB(185, 174, 152)
        blk.HorizontalAlignment = -4108             ' xlCenter
        blk.VerticalAlignment = -4108
        ws.Range(ws.Cells(R0, 3), ws.Cells(lastR, 4)).NumberFormat = "0.0000"
        ReDim nums(1 To n, 1 To 1)
        For i = 1 To n: nums(i, 1) = i: Next i
        ws.Range(ws.Cells(R0, 1), ws.Cells(lastR, 1)).Value = nums
        If specOK Then
            ' R1C1 formulas are identical for every row, so one assignment fills the whole column correctly
            ws.Range(ws.Cells(R0, 5), ws.Cells(lastR, 5)).FormulaR1C1 = "=IF(OR(RC3<R6C2,RC3>R6C3),""" & T_OOS() & """,""OK"")"
            ws.Range(ws.Cells(R0, 6), ws.Cells(lastR, 6)).FormulaR1C1 = "=IF(OR(RC4<R7C2,RC4>R7C3),""" & T_OOS() & """,""OK"")"
            ws.Range(ws.Cells(R0, 7), ws.Cells(lastR, 7)).FormulaR1C1 = "=IF(AND(RC5=""OK"",RC6=""OK""),""OK"",""NG"")"
            ws.Range(ws.Cells(R0, 7), ws.Cells(lastR, 7)).Font.Bold = True
            ws.Range(ws.Cells(R0, 7), ws.Cells(lastR, 7)).Font.Color = RGB(46, 125, 50)
            vals = ws.Range(ws.Cells(R0, 3), ws.Cells(lastR, 4)).Value
            For i = 1 To n
                x = vals(i, 1): y = vals(i, 2)
                bx = False: by = False
                If IsNum(x) Then bx = (CDbl(x) < sx0 Or CDbl(x) > sx1)
                If IsNum(y) Then by = (CDbl(y) < sy0 Or CDbl(y) > sy1)
                r = R0 + i - 1
                If bx Then AddU badV, ws.Cells(r, 3): AddU badJ, ws.Cells(r, 5)
                If by Then AddU badV, ws.Cells(r, 4): AddU badJ, ws.Cells(r, 6)
                If bx Or by Then AddU badG, ws.Cells(r, 7): nBad = nBad + 1
            Next i
            If Not badV Is Nothing Then MarkBad badV, True
            If Not badJ Is Nothing Then MarkBad badJ, False
            If Not badG Is Nothing Then MarkBad badG, False
        Else
            ws.Range(ws.Cells(R0, 5), ws.Cells(lastR, 7)).Value = "-"
        End If
    End If

    ' 4) statistics
    Dim XR As String, YR As String
    If n > 0 Then
        XR = "C" & R0 & ":C" & lastR
        YR = "D" & R0 & ":D" & lastR
        ws.Range("B11").Formula = "=AVERAGE(" & XR & ")"
        ws.Range("C11").Formula = "=AVERAGE(" & YR & ")"
        ws.Range("B12").Formula = "=MAX(" & XR & ")"
        ws.Range("C12").Formula = "=MAX(" & YR & ")"
        ws.Range("B13").Formula = "=MIN(" & XR & ")"
        ws.Range("C13").Formula = "=MIN(" & YR & ")"
        ws.Range("B14").Formula = "=IF(COUNT(" & XR & ")>1,STDEV(" & XR & "),"""")"
        ws.Range("C14").Formula = "=IF(COUNT(" & YR & ")>1,STDEV(" & YR & "),"""")"
        ws.Range("B15").Formula = "=COUNT(" & XR & ")"
        If specOK Then
            ws.Range("B16").Formula = "=COUNTIF(G" & R0 & ":G" & lastR & ",""NG"")"
            ws.Range("B17").Formula = "=IF(B15=0,"""",1-B16/B15)"
        Else
            ws.Range("B16").Value = T_NOSPEC()
            ws.Range("B17").Value = ""
        End If
    Else
        ws.Range("B11:C14").ClearContents
        ws.Range("B15").Value = 0
        ws.Range("B16").Value = 0
        ws.Range("B17").Value = ""
    End If
    Application.Calculate

    ClearMark ws.Range("B11:C11")
    ClearMark ws.Range("B16")
    If n > 0 And specOK Then
        If IsNum(ws.Range("B11").Value) Then
            If ws.Range("B11").Value < sx0 Or ws.Range("B11").Value > sx1 Then MarkBad ws.Range("B11"), True
        End If
        If IsNum(ws.Range("C11").Value) Then
            If ws.Range("C11").Value < sy0 Or ws.Range("C11").Value > sy1 Then MarkBad ws.Range("C11"), True
        End If
        If nBad > 0 Then MarkBad ws.Range("B16"), False
    End If

    FitColumns ws

    ' 5) redraw spec box / points / average on every diagram image
    On Error Resume Next
    DrawAll ws
    On Error GoTo Fail

    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = evts
    MsgBox T_DONE() & n & T_BAD() & nBad, vbInformation, T_TITLE()
    Exit Sub
Fail:
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = evts
    MsgBox T_ERR() & Err.Description, vbExclamation, T_TITLE()
End Sub

' redraw the diagram overlays only (used by the image selector and the check cells)
Public Sub RefreshChart(Optional ByVal mainOnly As Boolean = False)
    Dim ws As Worksheet, evts As Boolean
    Set ws = FindDataSheet()
    If ws Is Nothing Then Exit Sub
    evts = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error Resume Next
    DrawAll ws, mainOnly
    Application.ScreenUpdating = True
    Application.EnableEvents = evts
End Sub

' ---------- diagram overlays ----------
Private Sub DrawAll(ByVal ws As Worksheet, Optional ByVal mainOnly As Boolean = False)
    Dim ps As Worksheet
    Set ps = ParamSheet()
    If ps Is Nothing Then Exit Sub

    Dim specOK As Boolean, sx0 As Double, sx1 As Double, sy0 As Double, sy1 As Double
    specOK = IsNum(ws.Range("B6").Value) And IsNum(ws.Range("C6").Value) And IsNum(ws.Range("B7").Value) And IsNum(ws.Range("C7").Value)
    If specOK Then
        sx0 = CDbl(ws.Range("B6").Value): sx1 = CDbl(ws.Range("C6").Value)
        sy0 = CDbl(ws.Range("B7").Value): sy1 = CDbl(ws.Range("C7").Value)
        specOK = (sx0 < sx1 And sy0 < sy1)
    End If

    Dim lastR As Long, r As Long, n As Long
    Dim xs() As Double, ys() As Double
    lastR = LastRowIn(ws, 3, 4)
    ReDim xs(0 To 0): ReDim ys(0 To 0)
    If lastR >= R0 Then
        ReDim xs(1 To lastR - R0 + 1): ReDim ys(1 To lastR - R0 + 1)
        For r = R0 To lastR
            If IsNum(ws.Cells(r, 3).Value) And IsNum(ws.Cells(r, 4).Value) Then
                n = n + 1
                xs(n) = CDbl(ws.Cells(r, 3).Value)
                ys(n) = CDbl(ws.Cells(r, 4).Value)
            End If
        Next r
    End If

    mDot = Gp("dot", 0.0095)
    mBoxW = Gp("boxW", 3)
    mBoxCol = RGB(Gp("boxR", 255), Gp("boxG", 45), Gp("boxB", 85))
    On Error Resume Next
    ws.DisplayPageBreaks = False
    On Error GoTo 0

    Dim showBox As Boolean, showPts As Boolean, showAvg As Boolean, sel As String
    showBox = IsOn(ws.Range(C_BOX))
    showPts = IsOn(ws.Range(C_PTS))
    showAvg = IsOn(ws.Range(C_AVG))
    sel = CStr(ws.Range(C_IMG).Value)

    ' remove previous overlays everywhere
    Dim s As Worksheet, i As Long
    For Each s In ThisWorkbook.Worksheets
        For i = s.Shapes.Count To 1 Step -1
            If Left$(s.Shapes(i).Name, IIf(mainOnly, 8, 6)) = IIf(mainOnly, "CIEov_M_", "CIEov_") Then s.Shapes(i).Delete
        Next i
    Next s

    ' is the selected label one of the images?  otherwise fall back to the first one
    Dim k As Long, found As Boolean
    For k = 2 To 60
        If CStr(ps.Cells(k, 1).Value) = "" Then Exit For
        If CStr(ps.Cells(k, 2).Value) = sel Then found = True
    Next k
    If Not found Then
        sel = CStr(ps.Cells(2, 2).Value)
        ws.Range(C_IMG).Value = sel
    End If

    Dim nm As String, lbl As String, img As Shape
    On Error Resume Next
    For k = 2 To 60
        nm = CStr(ps.Cells(k, 1).Value)
        If nm = "" Then Exit For
        lbl = CStr(ps.Cells(k, 2).Value)
        For Each s In ThisWorkbook.Worksheets
            Set img = Nothing
            Set img = s.Shapes("M_" & nm)
            If Not img Is Nothing Then
                img.Left = s.Columns(9).Left + 2
                If lbl = sel Then
                    img.Visible = -1
                    Overlay s, img, "CIEov_M_" & nm & "_", ps, k, showBox, showPts, showAvg, specOK, sx0, sx1, sy0, sy1, xs, ys, n
                Else
                    img.Visible = 0
                End If
            End If
            Set img = Nothing
            If Not mainOnly Then Set img = s.Shapes("G_" & nm)
            If Not img Is Nothing Then
                Overlay s, img, "CIEov_G_" & nm & "_", ps, k, showBox, showPts, showAvg, specOK, sx0, sx1, sy0, sy1, xs, ys, n
            End If
        Next s
    Next k
End Sub

Private Function PX(ByVal x As Double) As Double
    PX = mIL + (mFL + (x - mVX0) / (mVX1 - mVX0) * mFW) * mIW
End Function
Private Function PY(ByVal y As Double) As Double
    PY = mIT + (mFT + (1 - (y - mVY0) / (mVY1 - mVY0)) * mFH) * mIH
End Function
Private Function InView(ByVal x As Double, ByVal y As Double) As Boolean
    InView = (x >= mVX0 And x <= mVX1 And y >= mVY0 And y <= mVY1)
End Function

Private Sub Track(ByVal s As Shape)
    If s Is Nothing Then Exit Sub
    mCnt = mCnt + 1
    If mCnt = 1 Then
        ReDim mNames(1 To 64)
    ElseIf mCnt > UBound(mNames) Then
        ReDim Preserve mNames(1 To mCnt * 2)
    End If
    mNames(mCnt) = s.Name
End Sub

Private Sub Overlay(ByVal sh As Worksheet, ByVal img As Shape, ByVal pre As String, ByVal ps As Worksheet, ByVal k As Long, _
                    ByVal showBox As Boolean, ByVal showPts As Boolean, ByVal showAvg As Boolean, ByVal specOK As Boolean, _
                    ByVal sx0 As Double, ByVal sx1 As Double, ByVal sy0 As Double, ByVal sy1 As Double, _
                    xs() As Double, ys() As Double, ByVal n As Long)
    On Error Resume Next
    mVX0 = ps.Cells(k, 3).Value: mVX1 = ps.Cells(k, 4).Value
    mVY0 = ps.Cells(k, 5).Value: mVY1 = ps.Cells(k, 6).Value
    mFL = ps.Cells(k, 7).Value: mFT = ps.Cells(k, 8).Value
    mFW = ps.Cells(k, 9).Value: mFH = ps.Cells(k, 10).Value
    mIL = img.Left: mIT = img.Top: mIW = img.Width: mIH = img.Height
    If mVX1 <= mVX0 Or mVY1 <= mVY0 Then Exit Sub
    mCnt = 0

    ' spec box: four edges, each clipped to the visible range
    If showBox And specOK Then
        Dim lw As Double, col As Long
        lw = mBoxW * mIW / 1000
        col = mBoxCol
        Edge sh, pre & "b1", sx0, sy0, sx1, sy0, lw, col
        Edge sh, pre & "b2", sx0, sy1, sx1, sy1, lw, col
        Edge sh, pre & "b3", sx0, sy0, sx0, sy1, lw, col
        Edge sh, pre & "b4", sx1, sy0, sx1, sy1, lw, col
    End If

    Dim d As Double, i As Long, bad As Boolean, sumX As Double, sumY As Double, s As Shape
    d = mIW * mDot
    For i = 1 To n
        sumX = sumX + xs(i): sumY = sumY + ys(i)
        If showPts And InView(xs(i), ys(i)) Then
            bad = False
            If specOK Then bad = (xs(i) < sx0 Or xs(i) > sx1 Or ys(i) < sy0 Or ys(i) > sy1)
            Set s = Nothing
            Set s = sh.Shapes.AddShape(9, PX(xs(i)) - d / 2, PY(ys(i)) - d / 2, d, d)      ' msoShapeOval
            If Not s Is Nothing Then
                s.Name = pre & "p" & i
                If bad Then
                    s.Fill.ForeColor.RGB = RGB(255, 59, 48): s.Line.ForeColor.RGB = RGB(255, 255, 255)
                Else
                    s.Fill.ForeColor.RGB = RGB(255, 255, 255): s.Line.ForeColor.RGB = RGB(0, 0, 0)
                End If
                s.Line.Weight = 0.75
                s.Placement = 3   ' xlFreeFloating
                Track s
            End If
        End If
    Next i

    If showAvg And n > 0 Then
        Dim ax As Double, ay As Double
        ax = sumX / n: ay = sumY / n
        If InView(ax, ay) Then
            bad = False
            If specOK Then bad = (ax < sx0 Or ax > sx1 Or ay < sy0 Or ay > sy1)
            Set s = Nothing
            Set s = sh.Shapes.AddShape(4, PX(ax) - d, PY(ay) - d, 2 * d, 2 * d)       ' msoShapeDiamond
            If s Is Nothing Then Set s = sh.Shapes.AddShape(9, PX(ax) - d, PY(ay) - d, 2 * d, 2 * d)
            If Not s Is Nothing Then
                s.Name = pre & "avg"
                If bad Then s.Fill.ForeColor.RGB = RGB(255, 59, 48) Else s.Fill.ForeColor.RGB = RGB(255, 214, 10)
                s.Line.ForeColor.RGB = RGB(0, 0, 0)
                s.Line.Weight = 1
                s.Placement = 3
                Track s
            End If
        End If
    End If

    ' group this image's overlay shapes: the next redraw deletes one group instead of many shapes
    If mCnt >= 2 Then
        Dim arr() As Variant, j As Long, g As Shape
        ReDim arr(1 To mCnt)
        For j = 1 To mCnt: arr(j) = mNames(j): Next j
        Set g = Nothing
        Set g = sh.Shapes.Range(arr).Group
        If Not g Is Nothing Then
            g.Name = pre & "grp"
            g.Placement = 3
        End If
    End If
End Sub

' one spec-box edge (horizontal or vertical), clipped to the view; dark outline + coloured line
Private Sub Edge(ByVal sh As Worksheet, ByVal nm As String, ByVal x0 As Double, ByVal y0 As Double, _
                 ByVal x1 As Double, ByVal y1 As Double, ByVal lw As Double, ByVal col As Long)
    Dim a As Double, b As Double, ext0 As Double, ext1 As Double, hw As Double
    Dim p0x As Double, p0y As Double, p1x As Double, p1y As Double, s As Shape
    hw = (lw + 2.6 * mIW / 1000) / 2
    If y0 = y1 Then
        If y0 < mVY0 Or y0 > mVY1 Then Exit Sub
        a = x0: b = x1
        If a < mVX0 Then a = mVX0
        If b > mVX1 Then b = mVX1
        If a >= b Then Exit Sub
        If a = x0 Then ext0 = hw
        If b = x1 Then ext1 = hw
        p0x = PX(a) - ext0: p1x = PX(b) + ext1: p0y = PY(y0): p1y = p0y
    Else
        If x0 < mVX0 Or x0 > mVX1 Then Exit Sub
        a = y0: b = y1
        If a < mVY0 Then a = mVY0
        If b > mVY1 Then b = mVY1
        If a >= b Then Exit Sub
        If a = y0 Then ext0 = hw
        If b = y1 Then ext1 = hw
        p0x = PX(x0): p1x = p0x: p0y = PY(a) + ext0: p1y = PY(b) - ext1
    End If
    Set s = Nothing
    Set s = sh.Shapes.AddLine(p0x, p0y, p1x, p1y)
    If Not s Is Nothing Then
        s.Name = nm & "o": s.Line.ForeColor.RGB = RGB(0, 0, 0): s.Line.Weight = lw + 2.6 * mIW / 1000: s.Placement = 3
        Track s
    End If
    Set s = Nothing
    Set s = sh.Shapes.AddLine(p0x, p0y, p1x, p1y)
    If Not s Is Nothing Then
        s.Name = nm & "c": s.Line.ForeColor.RGB = col: s.Line.Weight = lw: s.Placement = 3
        Track s
    End If
End Sub
