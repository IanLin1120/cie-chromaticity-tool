Attribute VB_Name = "Module1"
' CIE 1931 xy measurement report - recalculation macro
' Source is ASCII-only on purpose (Chinese text is built with ChrW so it works with any code page).
Option Explicit

Private Const R0 As Long = 21   ' first data row (row 20 = header)

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

Private Function Prm(ByVal key As String, ByVal dflt As Double) As Double
    Dim ps As Worksheet, r As Long
    Prm = dflt
    Set ps = ParamSheet()
    If ps Is Nothing Then Exit Function
    For r = 1 To 100
        If CStr(ps.Cells(r, 1).Value) = key Then
            If IsNum(ps.Cells(r, 2).Value) Then Prm = CDbl(ps.Cells(r, 2).Value)
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

Private Sub ClearMark(ByVal rg As Range)
    rg.Font.Color = RGB(0, 0, 0)
    rg.Font.Bold = False
    rg.Interior.Pattern = -4142   ' xlNone
End Sub

' ---------- main: assigned to the button ----------
Public Sub RecalcAll()
    Dim ws As Worksheet
    Set ws = FindDataSheet()
    If ws Is Nothing Then
        MsgBox T_NOSHEET(), vbExclamation, T_TITLE()
        Exit Sub
    End If

    Dim calcMode As Long
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = -4135   ' xlCalculationManual
    On Error GoTo Fail

    Dim r As Long, lastR As Long, lastAny As Long, n As Long, nBad As Long

    ' 1) remove rows that no longer have x and y (deleted / cleared measurements)
    lastAny = LastRowIn(ws, 1, 7)
    For r = lastAny To R0 Step -1
        If Trim$(CStr(ws.Cells(r, 3).Value)) = "" And Trim$(CStr(ws.Cells(r, 4).Value)) = "" Then ws.Rows(r).Delete
    Next r
    lastR = LastRowIn(ws, 3, 4)
    If lastR < R0 Then n = 0 Else n = lastR - R0 + 1

    ' 2) spec limits
    Dim sx0 As Variant, sx1 As Variant, sy0 As Variant, sy1 As Variant, specOK As Boolean
    sx0 = ws.Range("B6").Value: sx1 = ws.Range("C6").Value
    sy0 = ws.Range("B7").Value: sy1 = ws.Range("C7").Value
    specOK = IsNum(sx0) And IsNum(sx1) And IsNum(sy0) And IsNum(sy1)

    ' 3) numbering, judgement formulas and highlighting
    Dim x As Variant, y As Variant, bx As Boolean, by As Boolean, rowRg As Range
    For r = R0 To lastR
        ws.Cells(r, 1).Value = r - R0 + 1
        x = ws.Cells(r, 3).Value: y = ws.Cells(r, 4).Value
        bx = False: by = False
        If specOK Then
            If IsNum(x) Then bx = (CDbl(x) < sx0 Or CDbl(x) > sx1)
            If IsNum(y) Then by = (CDbl(y) < sy0 Or CDbl(y) > sy1)
            ws.Cells(r, 5).Formula = "=IF(OR(C" & r & "<$B$6,C" & r & ">$C$6),""" & T_OOS() & """,""OK"")"
            ws.Cells(r, 6).Formula = "=IF(OR(D" & r & "<$B$7,D" & r & ">$C$7),""" & T_OOS() & """,""OK"")"
            ws.Cells(r, 7).Formula = "=IF(AND(E" & r & "=""OK"",F" & r & "=""OK""),""OK"",""NG"")"
        Else
            ws.Range(ws.Cells(r, 5), ws.Cells(r, 7)).Value = "-"
        End If
        Set rowRg = ws.Range(ws.Cells(r, 1), ws.Cells(r, 7))
        ClearMark rowRg
        rowRg.Borders.LineStyle = 1                 ' xlContinuous
        rowRg.Borders.Color = RGB(185, 174, 152)
        rowRg.HorizontalAlignment = -4108           ' xlCenter
        rowRg.VerticalAlignment = -4108
        ws.Range(ws.Cells(r, 3), ws.Cells(r, 4)).NumberFormat = "0.0000"
        If bx Then MarkBad ws.Cells(r, 3), True: MarkBad ws.Cells(r, 5), False
        If by Then MarkBad ws.Cells(r, 4), True: MarkBad ws.Cells(r, 6), False
        If specOK Then
            If bx Or by Then
                MarkBad ws.Cells(r, 7), False
                nBad = nBad + 1
            Else
                ws.Cells(r, 7).Font.Bold = True
                ws.Cells(r, 7).Font.Color = RGB(46, 125, 50)
            End If
        End If
    Next r

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

    ' 5) redraw the measurement points on the diagram
    On Error Resume Next
    DrawPoints ws, lastR, specOK, sx0, sx1, sy0, sy1
    On Error GoTo Fail

    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    MsgBox T_DONE() & n & T_BAD() & nBad, vbInformation, T_TITLE()
    Exit Sub
Fail:
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    MsgBox T_ERR() & Err.Description, vbExclamation, T_TITLE()
End Sub

' ---------- diagram: points are drawn as shapes on top of the bImg image ----------
Private Sub DrawPoints(ByVal ws As Worksheet, ByVal lastR As Long, ByVal specOK As Boolean, _
                       ByVal sx0 As Variant, ByVal sx1 As Variant, ByVal sy0 As Variant, ByVal sy1 As Variant)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 6) = "CIEpt_" Then ws.Shapes(i).Delete
    Next i

    Dim img As Shape, bImg As Shape
    Set img = Nothing: Set bImg = Nothing
    On Error Resume Next
    Set img = ws.Shapes("CIE_Chart")
    Set bImg = ws.Shapes("CIE_Base")
    On Error GoTo 0
    If bImg Is Nothing Then Exit Sub
    If Not img Is Nothing Then
        bImg.LockAspectRatio = 0
        bImg.Left = img.Left: bImg.Top = img.Top
        bImg.Width = img.Width: bImg.Height = img.Height
        img.Visible = 0
    End If
    bImg.Visible = -1

    Dim vx0 As Double, vx1 As Double, vy0 As Double, vy1 As Double
    Dim fl As Double, ft As Double, fw As Double, fh As Double, d As Double
    vx0 = Prm("vx0", 0): vx1 = Prm("vx1", 0.8): vy0 = Prm("vy0", 0): vy1 = Prm("vy1", 0.9)
    fl = Prm("fl", 0.1): ft = Prm("ft", 0.07): fw = Prm("fw", 0.86): fh = Prm("fh", 0.84)
    d = bImg.Width * Prm("dot", 0.011)
    Dim showPts As Boolean, showAvg As Boolean
    showPts = (Prm("pts", 1) <> 0)
    showAvg = (Prm("avg", 1) <> 0)

    Dim r As Long, x As Double, y As Double, bad As Boolean, px As Double, py As Double
    Dim sumX As Double, sumY As Double, cnt As Long, s As Shape
    On Error Resume Next      ' a styling property that is not supported must not stop the drawing
    For r = R0 To lastR
        If IsNum(ws.Cells(r, 3).Value) And IsNum(ws.Cells(r, 4).Value) Then
            x = CDbl(ws.Cells(r, 3).Value): y = CDbl(ws.Cells(r, 4).Value)
            sumX = sumX + x: sumY = sumY + y: cnt = cnt + 1
            If showPts And x >= vx0 And x <= vx1 And y >= vy0 And y <= vy1 Then
                bad = False
                If specOK Then bad = (x < sx0 Or x > sx1 Or y < sy0 Or y > sy1)
                px = bImg.Left + (fl + (x - vx0) / (vx1 - vx0) * fw) * bImg.Width
                py = bImg.Top + (ft + (1 - (y - vy0) / (vy1 - vy0)) * fh) * bImg.Height
                Set s = Nothing
                Set s = ws.Shapes.AddShape(9, px - d / 2, py - d / 2, d, d)      ' msoShapeOval
                If Not s Is Nothing Then
                    s.Name = "CIEpt_" & r
                    If bad Then
                        s.Fill.ForeColor.RGB = RGB(255, 59, 48): s.Line.ForeColor.RGB = RGB(255, 255, 255)
                    Else
                        s.Fill.ForeColor.RGB = RGB(255, 255, 255): s.Line.ForeColor.RGB = RGB(0, 0, 0)
                    End If
                    s.Line.Weight = 0.75
                    s.Placement = 2   ' xlMove
                End If
            End If
        End If
    Next r

    If showAvg And cnt > 0 Then
        x = sumX / cnt: y = sumY / cnt
        If x >= vx0 And x <= vx1 And y >= vy0 And y <= vy1 Then
            bad = False
            If specOK Then bad = (x < sx0 Or x > sx1 Or y < sy0 Or y > sy1)
            px = bImg.Left + (fl + (x - vx0) / (vx1 - vx0) * fw) * bImg.Width
            py = bImg.Top + (ft + (1 - (y - vy0) / (vy1 - vy0)) * fh) * bImg.Height
            Set s = Nothing
            Set s = ws.Shapes.AddShape(4, px - d, py - d, 2 * d, 2 * d)       ' msoShapeDiamond
            If s Is Nothing Then Set s = ws.Shapes.AddShape(9, px - d, py - d, 2 * d, 2 * d)
            If Not s Is Nothing Then
                s.Name = "CIEpt_avg"
                If bad Then s.Fill.ForeColor.RGB = RGB(255, 59, 48) Else s.Fill.ForeColor.RGB = RGB(255, 214, 10)
                s.Line.ForeColor.RGB = RGB(0, 0, 0)
                s.Line.Weight = 1
                s.Placement = 2
            End If
        End If
    End If
End Sub
