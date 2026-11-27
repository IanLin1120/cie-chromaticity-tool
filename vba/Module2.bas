Attribute VB_Name = "Module2"
' Luminance sheet ("A19 = No."): recalculation + histogram / trend charts redrawn with shapes.
' Runs from the button (RecalcLum) and automatically when L values or the spec row change (ThisWorkbook event).
' Source is ASCII-only on purpose (Chinese text is built with ChrW; chart labels come from the parameter sheet).
Option Explicit

Private Const LR0 As Long = 20                 ' first data row (row 19 = header)
Private Const LFONT As String = "Microsoft JhengHei"
Private Const LMONO As String = "Consolas"

' data (filled by ReadLum)
Private mLs() As Double, mSeq() As Double, mBad() As Boolean, mRow() As Long, mN As Long
Private mLo As Double, mHi As Double, mTy As Double, mHLo As Boolean, mHHi As Boolean, mHTy As Boolean
Private mAvg As Double, mSd As Double, mMax As Double, mMin As Double, mBelow As Long, mAbove As Long
Private mCpk As Double, mHasCpk As Boolean, mOkSd As Boolean

' chart settings (from the parameter sheet)
Private mBins As String, mNormal As Boolean, mMeanOn As Boolean, mTypOn As Boolean, mStatOn As Boolean, mIncSpec As Boolean

' drawing state
Private mSh As Worksheet
Private mU As Double, mL As Double, mT As Double, mPW As Double, mPH As Double, mW As Double, mH As Double
Private mOX As Double, mOY As Double
Private mX0 As Double, mX1 As Double, mY0 As Double, mY1 As Double
Private mNm() As String, mCnt As Long, mPre As String

' ---------- text (Traditional Chinese via ChrW) ----------
Private Function T_LTITLE() As String
    T_LTITLE = {{亮度量測}}
End Function
Private Function T_LDONE() As String
    T_LDONE = {{亮度已重新計算完成。筆數：}}
End Function
Private Function T_LBAD() As String
    T_LBAD = {{，不合格筆數：}}
End Function
Private Function T_LERR() As String
    T_LERR = {{亮度重新計算時發生錯誤：}}
End Function
Private Function T_LNOSHEET() As String
    T_LNOSHEET = {{找不到亮度量測工作表（A19 應為 No.）}}
End Function
Private Function T_STATLABEL(ByVal i As Long) As String
    Select Case i
        Case 0: T_STATLABEL = {{筆數}}
        Case 1: T_STATLABEL = {{平均}}
        Case 2: T_STATLABEL = {{最大}}
        Case 3: T_STATLABEL = {{最小}}
        Case 4: T_STATLABEL = {{標準差}}
        Case 6: T_STATLABEL = {{不合格筆數}}
        Case 7: T_STATLABEL = {{合格率}}
    End Select
End Function

' ---------- utilities ----------
Private Function IsNumV(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Trim$(v) = "" Then Exit Function
    End If
    IsNumV = IsNumeric(v)
End Function

Private Function LParamSheet() As Worksheet
    Dim s As Worksheet
    For Each s In ThisWorkbook.Worksheets
        If s.Name = "_cie_params" Then
            Set LParamSheet = s
            Exit Function
        End If
    Next s
End Function

' luminance settings: key in column O, value in column P of the parameter sheet
Private Function Lp(ByVal key As String, ByVal dflt As String) As String
    Dim ps As Worksheet, r As Long
    Lp = dflt
    Set ps = LParamSheet()
    If ps Is Nothing Then Exit Function
    For r = 1 To 80
        If CStr(ps.Cells(r, 15).Value) = key Then
            Lp = CStr(ps.Cells(r, 16).Value)
            Exit Function
        End If
    Next r
End Function

Public Function IsLumSheet(ByVal s As Object) As Boolean
    On Error Resume Next
    IsLumSheet = (CStr(s.Range("A19").Value) = "No.")
End Function

Private Function FindLumSheet() As Worksheet
    Dim s As Worksheet
    If IsLumSheet(ActiveSheet) Then
        Set FindLumSheet = ActiveSheet
        Exit Function
    End If
    For Each s In ThisWorkbook.Worksheets
        If IsLumSheet(s) Then
            Set FindLumSheet = s
            Exit Function
        End If
    Next s
End Function

Private Function LastLumRow(ByVal ws As Worksheet) As Long
    Dim c As Long, r As Long
    LastLumRow = LR0 - 1
    For c = 1 To 4
        r = ws.Cells(ws.Rows.Count, c).End(-4162).Row   ' xlUp
        If r > LastLumRow Then LastLumRow = r
    Next c
End Function

Private Sub LMark(ByVal rg As Range, ByVal withFill As Boolean)
    rg.Font.Color = RGB(255, 0, 0)
    rg.Font.Bold = True
    If withFill Then rg.Interior.Color = RGB(253, 226, 225)
End Sub

Private Sub LClear(ByVal rg As Range)
    rg.Font.Color = RGB(0, 0, 0)
    rg.Font.Bold = False
    rg.Interior.Pattern = -4142   ' xlNone
End Sub

Private Sub LBorder(ByVal rg As Range)
    rg.Borders.LineStyle = 1      ' xlContinuous
    rg.Borders.Color = RGB(185, 174, 152)
    rg.HorizontalAlignment = -4108
    rg.VerticalAlignment = -4108
    rg.Font.Name = LFONT
End Sub

Private Sub AddR(ByRef u As Range, ByVal c As Range)
    If u Is Nothing Then
        Set u = c
    Else
        Set u = Application.Union(u, c)
    End If
End Sub

' ---------- data ----------
Private Sub ReadLum(ByVal ws As Worksheet)
    Dim v As Variant
    mHLo = False: mHHi = False: mHTy = False
    v = ws.Range("B6").Value: If IsNumV(v) Then mLo = CDbl(v): mHLo = True
    v = ws.Range("D6").Value: If IsNumV(v) Then mHi = CDbl(v): mHHi = True
    v = ws.Range("C6").Value: If IsNumV(v) Then mTy = CDbl(v): mHTy = True
    If mHLo And mHHi Then
        If mLo >= mHi Then mHLo = False: mHHi = False
    End If

    Dim lastR As Long, vals As Variant, i As Long, k As Long
    mN = 0: mBelow = 0: mAbove = 0
    ReDim mLs(0 To 0): ReDim mSeq(0 To 0): ReDim mBad(0 To 0): ReDim mRow(0 To 0)
    lastR = LastLumRow(ws)
    If lastR >= LR0 Then
        vals = ws.Range(ws.Cells(LR0, 1), ws.Cells(lastR, 4)).Value
        k = lastR - LR0 + 1
        ReDim mLs(1 To k): ReDim mSeq(1 To k): ReDim mBad(1 To k): ReDim mRow(1 To k)
        For i = 1 To k
            If IsNumV(vals(i, 3)) Then
                mN = mN + 1
                mLs(mN) = CDbl(vals(i, 3))
                If IsNumV(vals(i, 1)) Then mSeq(mN) = CDbl(vals(i, 1)) Else mSeq(mN) = mN
                mRow(mN) = LR0 + i - 1
                mBad(mN) = False
                If mHLo Then
                    If mLs(mN) < mLo Then mBad(mN) = True: mBelow = mBelow + 1
                End If
                If mHHi Then
                    If mLs(mN) > mHi Then mBad(mN) = True: mAbove = mAbove + 1
                End If
            End If
        Next i
    End If

    Dim s As Double
    mAvg = 0: mSd = 0: mMax = 0: mMin = 0: mOkSd = False: mHasCpk = False
    If mN = 0 Then Exit Sub
    mMax = mLs(1): mMin = mLs(1)
    For i = 1 To mN
        s = s + mLs(i)
        If mLs(i) > mMax Then mMax = mLs(i)
        If mLs(i) < mMin Then mMin = mLs(i)
    Next i
    mAvg = s / mN
    If mN > 1 Then
        s = 0
        For i = 1 To mN: s = s + (mLs(i) - mAvg) ^ 2: Next i
        mSd = Sqr(s / (mN - 1))
        mOkSd = (mSd > 0)
    End If
    If mOkSd And (mHLo Or mHHi) Then
        mHasCpk = True
        If mHLo Then mCpk = (mAvg - mLo) / (3 * mSd)
        If mHHi Then
            If Not mHLo Then
                mCpk = (mHi - mAvg) / (3 * mSd)
            ElseIf (mHi - mAvg) / (3 * mSd) < mCpk Then
                mCpk = (mHi - mAvg) / (3 * mSd)
            End If
        End If
    End If
End Sub

Private Function CpkName() As String
    If mHLo And mHHi Then
        CpkName = "Cpk"
    ElseIf mHLo Then
        CpkName = "Cpl"
    Else
        CpkName = "Cpu"
    End If
End Function

' ---------- sheet: numbering gaps, judgement, statistics, highlighting ----------
Private Sub LumUpdate(ByVal ws As Worksheet)
    Dim lastR As Long, i As Long, r As Long, vals As Variant, mx As Double, has As Boolean
    lastR = LastLumRow(ws)

    ' fill empty sample numbers (column A) for newly added rows
    If lastR >= LR0 Then
        vals = ws.Range(ws.Cells(LR0, 1), ws.Cells(lastR, 3)).Value
        For i = 1 To lastR - LR0 + 1
            If IsNumV(vals(i, 1)) Then
                If CDbl(vals(i, 1)) > mx Then mx = CDbl(vals(i, 1))
            End If
        Next i
        For i = 1 To lastR - LR0 + 1
            If IsEmpty(vals(i, 1)) And IsNumV(vals(i, 3)) Then
                mx = mx + 1
                ws.Cells(LR0 + i - 1, 1).Value = mx
            End If
        Next i
    End If

    ReadLum ws
    has = mHLo Or mHHi

    ' judgement column D + highlighting
    Dim cond As String
    If lastR >= LR0 Then
        Dim blk As Range
        Set blk = ws.Range(ws.Cells(LR0, 1), ws.Cells(lastR, 4))
        LClear blk
        LBorder blk
        If has Then
            If mHLo Then cond = "RC3<R6C2"
            If mHHi Then
                If cond <> "" Then cond = cond & ","
                cond = cond & "RC3>R6C4"
            End If
            ws.Range(ws.Cells(LR0, 4), ws.Cells(lastR, 4)).FormulaR1C1 = "=IF(RC3="""","""",IF(OR(" & cond & "),""NG"",""OK""))"
            ws.Range(ws.Cells(LR0, 4), ws.Cells(lastR, 4)).Font.Bold = True
            ws.Range(ws.Cells(LR0, 4), ws.Cells(lastR, 4)).Font.Color = RGB(46, 125, 50)
            For i = 1 To mN
                If mBad(i) Then
                    LMark ws.Cells(mRow(i), 3), True
                    LMark ws.Cells(mRow(i), 4), False
                End If
            Next i
        Else
            ws.Range(ws.Cells(LR0, 4), ws.Cells(lastR, 4)).Value = ChrW(&H2014)
        End If
    End If

    ' statistics block (rows 9..16), always rebuilt so it follows the current data range and spec
    Dim RG As String, labs As Variant, fmts As Variant, fs(0 To 7) As String, nRows As Long
    If lastR >= LR0 Then RG = "C" & LR0 & ":C" & lastR Else RG = "C" & LR0 & ":C" & LR0
    fs(0) = "=COUNT(" & RG & ")"
    fs(1) = "=AVERAGE(" & RG & ")"
    fs(2) = "=MAX(" & RG & ")"
    fs(3) = "=MIN(" & RG & ")"
    fs(4) = "=IF(COUNT(" & RG & ")>1,STDEV(" & RG & "),"""")"
    If mHLo And mHHi Then
        fs(5) = "=IF(OR(B13="""",B13=0),"""",MIN((B10-B6)/(3*B13),(D6-B10)/(3*B13)))"
    ElseIf mHLo Then
        fs(5) = "=IF(OR(B13="""",B13=0),"""",(B10-B6)/(3*B13))"
    Else
        fs(5) = "=IF(OR(B13="""",B13=0),"""",(D6-B10)/(3*B13))"
    End If
    fs(6) = "=COUNTIF(D" & LR0 & ":D" & IIf(lastR >= LR0, lastR, LR0) & ",""NG"")"
    fs(7) = "=IF(B9=0,"""",1-B15/B9)"
    fmts = Array("0", "0.0", "0.0", "0.0", "0.00", "0.00", "0", "0.0%")
    If has Then nRows = 8 Else nRows = 5
    For i = 0 To 7
        r = 9 + i
        If i < nRows Then
            ws.Range("B" & r & ":C" & r).Merge
            If i = 5 Then ws.Range("A" & r).Value = CpkName() Else ws.Range("A" & r).Value = T_STATLABEL(i)
            ws.Range("A" & r).Font.Bold = True
            ws.Range("B" & r).Formula = fs(i)
            ws.Range("B" & r).NumberFormat = fmts(i)
            LClear ws.Range("B" & r & ":C" & r)
            LBorder ws.Range("A" & r & ":C" & r)
        Else
            ws.Range("A" & r & ":C" & r).ClearContents
            ws.Range("A" & r & ":C" & r).Borders.LineStyle = -4142
            LClear ws.Range("A" & r & ":C" & r)
        End If
    Next i
    If has Then
        If mHasCpk Then
            If mCpk < 1.33 Then LMark ws.Range("B14:C14"), True
        End If
        If mBelow + mAbove > 0 Then
            LMark ws.Range("B15:C15"), True
            LMark ws.Range("B16:C16"), True
        End If
    End If
End Sub

' ---------- entry points ----------
Public Sub RecalcLum()
    Dim ws As Worksheet
    Set ws = FindLumSheet()
    If ws Is Nothing Then
        MsgBox T_LNOSHEET(), vbExclamation, T_LTITLE()
        Exit Sub
    End If
    Dim evts As Boolean, calcMode As Long
    evts = Application.EnableEvents
    calcMode = Application.Calculation
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = -4135   ' xlCalculationManual
    On Error GoTo Fail

    ' remove rows whose luminance cell was cleared
    Dim r As Long, delRg As Range
    For r = LastLumRow(ws) To LR0 Step -1
        If Trim$(CStr(ws.Cells(r, 3).Value)) = "" Then AddR delRg, ws.Rows(r)
    Next r
    If Not delRg Is Nothing Then delRg.EntireRow.Delete

    LumUpdate ws
    Application.Calculate
    On Error Resume Next
    LumDraw ws
    On Error GoTo Fail

    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = evts
    MsgBox T_LDONE() & mN & T_LBAD() & (mBelow + mAbove), vbInformation, T_LTITLE()
    Exit Sub
Fail:
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = evts
    MsgBox T_LERR() & Err.Description, vbExclamation, T_LTITLE()
End Sub

' called by ThisWorkbook when a luminance value or the spec row changes: no message, no row deletion
Public Sub LumAuto(ByVal s As Object)
    Dim ws As Worksheet, evts As Boolean, calcMode As Long
    Set ws = s
    evts = Application.EnableEvents
    calcMode = Application.Calculation
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = -4135
    On Error Resume Next
    LumUpdate ws
    Application.Calculate
    LumDraw ws
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = evts
End Sub

' ---------- charts ----------
Private Sub LumDraw(ByVal ws As Worksheet)
    Set mSh = ws
    mBins = Lp("lum_bins", "auto")
    mNormal = (Lp("lum_normal", "1") = "1")
    mMeanOn = (Lp("lum_mean", "1") = "1")
    mTypOn = (Lp("lum_typ", "1") = "1")
    mStatOn = (Lp("lum_stat", "1") = "1")
    mIncSpec = (Lp("lum_incSpec", "0") = "1")

    ' keep the position / width of the existing charts (exported pictures or earlier groups)
    Dim hx As Double, hy As Double, hw As Double, rx As Double, ry As Double, rw As Double
    Dim shp As Shape, i As Long
    hx = ws.Columns(6).Left: hy = ws.Rows(3).Top: hw = 480
    Set shp = Nothing
    On Error Resume Next
    Set shp = ws.Shapes("LUMchart_hist")
    On Error GoTo 0
    If Not shp Is Nothing Then hx = shp.Left: hy = shp.Top: hw = shp.Width
    rx = hx: ry = hy + hw * 0.64 + 18: rw = hw
    Set shp = Nothing
    On Error Resume Next
    Set shp = ws.Shapes("LUMchart_run")
    On Error GoTo 0
    If Not shp Is Nothing Then rx = shp.Left: ry = shp.Top: rw = shp.Width

    ' remove the old charts and any pieces left over from an interrupted redraw
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 9) = "LUMchart_" Then ws.Shapes(i).Delete
    Next i
    DrawChart False, "LUMchart_hist", hx, hy, hw
    DrawChart True, "LUMchart_run", rx, ry, rw
End Sub

Private Function NiceStep(ByVal raw As Double) As Double
    Dim p As Double, f As Double
    If raw <= 0 Then
        NiceStep = 1
        Exit Function
    End If
    p = 10 ^ Int(Log(raw) / Log(10#))
    f = raw / p
    If f < 1.5 Then
        NiceStep = p
    ElseIf f < 3.5 Then
        NiceStep = 2 * p
    ElseIf f < 7.5 Then
        NiceStep = 5 * p
    Else
        NiceStep = 10 * p
    End If
End Function

Private Function LumDec(ByVal s As Double) As Long
    If s >= 1 Then
        LumDec = 0
    Else
        LumDec = -Int(-(-Log(s) / Log(10#) - 0.000000001))
        If LumDec > 4 Then LumDec = 4
    End If
End Function

Private Function FmtL(ByVal v As Double) As String
    If Abs(v) >= 100 Then
        FmtL = Format$(v, "0.0")
    ElseIf Abs(v) >= 10 Then
        FmtL = Format$(v, "0.00")
    Else
        FmtL = Format$(v, "0.000")
    End If
End Function

Private Function FmtN(ByVal v As Double) As String
    FmtN = CStr(Round(v, 4))
End Function

Private Function FmtD(ByVal v As Double, ByVal d As Long) As String
    If d <= 0 Then
        FmtD = Format$(Round(v, 0), "0")
    Else
        FmtD = Format$(v, "0." & String$(d, "0"))
    End If
End Function

' approximate text width (CJK ~ 1 em, Latin / digits ~ 0.56 em)
Private Function TextW(ByVal t As String, ByVal sz As Double) As Double
    Dim i As Long, c As Long
    For i = 1 To Len(t)
        c = AscW(Mid$(t, i, 1))
        If c < 0 Then c = c + 65536
        If c >= &H2E80 Then TextW = TextW + sz Else TextW = TextW + 0.56 * sz
    Next i
End Function

Private Function XP(ByVal v As Double) As Double
    XP = mL + (v - mX0) / (mX1 - mX0) * mPW
End Function
Private Function YP(ByVal v As Double) As Double
    YP = mT + mPH - (v - mY0) / (mY1 - mY0) * mPH
End Function

Private Sub Keep(ByVal s As Shape)
    If s Is Nothing Then Exit Sub
    mCnt = mCnt + 1
    If mCnt = 1 Then
        ReDim mNm(1 To 256)
    ElseIf mCnt > UBound(mNm) Then
        ReDim Preserve mNm(1 To mCnt * 2)
    End If
    s.Name = mPre & mCnt
    mNm(mCnt) = s.Name
End Sub

' fillCol / lineCol < 0 = none; transparency 0..1
Private Sub DRect(ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                  ByVal fillCol As Long, ByVal fillT As Double, ByVal lineCol As Long, ByVal lineT As Double, ByVal lw As Double)
    Dim s As Shape
    If w <= 0 Or h <= 0 Then Exit Sub
    Set s = mSh.Shapes.AddShape(1, mOX + x, mOY + y, w, h)   ' msoShapeRectangle
    StyleShape s, fillCol, fillT, lineCol, lineT, lw
    Keep s
End Sub

Private Sub DOval(ByVal cx As Double, ByVal cy As Double, ByVal r As Double, ByVal fillCol As Long, ByVal lineCol As Long, ByVal lw As Double)
    Dim s As Shape
    Set s = mSh.Shapes.AddShape(9, mOX + cx - r, mOY + cy - r, 2 * r, 2 * r)   ' msoShapeOval
    StyleShape s, fillCol, 0, lineCol, 0, lw
    Keep s
End Sub

Private Sub StyleShape(ByVal s As Shape, ByVal fillCol As Long, ByVal fillT As Double, ByVal lineCol As Long, ByVal lineT As Double, ByVal lw As Double)
    On Error Resume Next
    s.Shadow.Visible = 0
    If fillCol < 0 Then
        s.Fill.Visible = 0
    Else
        s.Fill.Visible = -1
        s.Fill.Solid
        s.Fill.ForeColor.RGB = fillCol
        s.Fill.Transparency = fillT
    End If
    If lineCol < 0 Then
        s.Line.Visible = 0
    Else
        s.Line.Visible = -1
        s.Line.ForeColor.RGB = lineCol
        s.Line.Transparency = lineT
        s.Line.Weight = lw
    End If
End Sub

' dash: 0 = solid, 4 = msoLineDash, 2 = msoLineSquareDot
Private Sub DLine(ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double, _
                  ByVal col As Long, ByVal tr As Double, ByVal lw As Double, ByVal dash As Long)
    Dim s As Shape
    Set s = mSh.Shapes.AddLine(mOX + x0, mOY + y0, mOX + x1, mOY + y1)
    On Error Resume Next
    s.Line.ForeColor.RGB = col
    s.Line.Transparency = tr
    s.Line.Weight = lw
    If dash > 0 Then s.Line.DashStyle = dash
    On Error GoTo 0
    Keep s
End Sub

Private Sub DPoly(px() As Double, py() As Double, ByVal n As Long, ByVal col As Long, ByVal tr As Double, ByVal lw As Double)
    Dim arr() As Single, i As Long, s As Shape
    If n < 2 Then Exit Sub
    ReDim arr(1 To n, 1 To 2)
    For i = 1 To n
        arr(i, 1) = mOX + px(i)
        arr(i, 2) = mOY + py(i)
    Next i
    Set s = Nothing
    On Error Resume Next
    Set s = mSh.Shapes.AddPolyline(arr)
    If s Is Nothing Then
        ' no polyline support: draw the segments as lines
        On Error GoTo 0
        For i = 1 To n - 1
            DLine px(i), py(i), px(i + 1), py(i + 1), col, tr, lw, 0
        Next i
        Exit Sub
    End If
    s.Fill.Visible = 0
    s.Line.ForeColor.RGB = col
    s.Line.Transparency = tr
    s.Line.Weight = lw
    On Error GoTo 0
    Keep s
End Sub

' ha: 0 left, 1 center, 2 right (x is that edge); va: 0 top, 1 middle, 2 bottom (y is that edge)
Private Sub DText(ByVal t As String, ByVal x As Double, ByVal y As Double, ByVal sz As Double, ByVal bold As Boolean, _
                  ByVal col As Long, ByVal tr As Double, ByVal ha As Long, ByVal va As Long, ByVal mono As Boolean, ByVal rot As Boolean)
    Dim s As Shape, w As Double, h As Double, l As Double, tp As Double
    If t = "" Then Exit Sub
    w = TextW(t, sz) + sz * 1.2
    h = sz * 1.45
    If rot Then
        l = x - w / 2: tp = y - h / 2
    Else
        Select Case ha
            Case 0: l = x
            Case 1: l = x - w / 2
            Case Else: l = x - w
        End Select
        Select Case va
            Case 0: tp = y
            Case 1: tp = y - h / 2
            Case Else: tp = y - h
        End Select
    End If
    Set s = Nothing
    On Error Resume Next
    Set s = mSh.Shapes.AddTextbox(1, mOX + l, mOY + tp, w, h)   ' msoTextOrientationHorizontal
    If s Is Nothing Then Set s = mSh.Shapes.AddShape(1, mOX + l, mOY + tp, w, h)
    If s Is Nothing Then Exit Sub
    s.Fill.Visible = 0
    s.Line.Visible = 0
    With s.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = 0
        .AutoSize = 0
        If rot Then
            .VerticalAnchor = 3
        Else
            .VerticalAnchor = IIf(va = 0, 1, IIf(va = 1, 3, 4))   ' msoAnchorTop / Middle / Bottom
        End If
        .TextRange.Text = t
        .TextRange.ParagraphFormat.Alignment = IIf(rot, 2, IIf(ha = 0, 1, IIf(ha = 1, 2, 3)))
        .TextRange.Font.Size = sz
        .TextRange.Font.Bold = IIf(bold, -1, 0)
        .TextRange.Font.Name = IIf(mono, LMONO, LFONT)
        .TextRange.Font.NameFarEast = LFONT
        .TextRange.Font.Fill.ForeColor.RGB = col
        .TextRange.Font.Fill.Transparency = tr
    End With
    If s.TextFrame2.TextRange.Text <> t Then
        ' older object model (no TextFrame2): plain TextFrame
        With s.TextFrame
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            .Characters.Text = t
            .HorizontalAlignment = IIf(rot, -4108, IIf(ha = 0, -4131, IIf(ha = 1, -4108, -4152)))
            .VerticalAlignment = IIf(rot, -4108, IIf(va = 0, -4160, IIf(va = 1, -4108, -4107)))
            .Characters.Font.Size = sz
            .Characters.Font.Bold = bold
            .Characters.Font.Name = IIf(mono, LMONO, LFONT)
            .Characters.Font.Color = col
        End With
    End If
    If rot Then s.Rotation = 270
    On Error GoTo 0
    Keep s
End Sub

' tick values from a to b with step st (returned as a 1-based array, count in cnt)
Private Function TickList(ByVal a As Double, ByVal b As Double, ByVal st As Double, ByRef cnt As Long) As Variant
    Dim out() As Double, k As Long
    cnt = 0
    ReDim out(1 To 401)
    k = -Int(-(a / st - 0.0000001))
    Do While k * st <= b + st * 0.0000001 And cnt < 400
        cnt = cnt + 1
        out(cnt) = k * st
        k = k + 1
    Loop
    TickList = out
End Function

Private Sub DrawChart(ByVal isRun As Boolean, ByVal nm As String, ByVal fx As Double, ByVal fy As Double, ByVal fw As Double)
    Const C_OK As Long = 16755008      ' RGB(64,169,255)
    Dim cBad As Long, cMean As Long, cTyp As Long, cFit As Long, cWhite As Long, cTick As Long
    cBad = RGB(255, 69, 58): cMean = RGB(255, 214, 10): cTyp = RGB(48, 209, 88): cFit = RGB(255, 159, 10)
    cWhite = RGB(255, 255, 255): cTick = RGB(232, 236, 244)

    Dim panel As Boolean, u As Double, i As Long, j As Long
    mCnt = 0: mPre = nm & "_"
    mOX = fx: mOY = fy
    panel = mStatOn And mN > 0
    u = fw / 1000: mU = u
    mW = fw: mH = fw * 0.64
    mL = 108 * u: mT = 74 * u
    mPW = mW - mL - IIf(panel, 262, 36) * u
    mPH = mH - mT - 98 * u

    DRect 0, 0, mW, mH, RGB(0, 0, 0), 0, -1, 0, 0
    Dim tt As String
    If isRun Then tt = Lp("lum_t_run", "Luminance Trend") Else tt = Lp("lum_t_lum", "Luminance Distribution")
    DText tt, mL, mT - 26 * u, 27 * u, True, cWhite, 0, 0, 2, False, False
    If TextW(tt, 27 * u) + TextW(Lp("lum_lobs", ""), 16 * u) + 30 * u < mPW Then _
        DText Lp("lum_lobs", ""), mL + mPW, mT - 28 * u, 16 * u, False, cWhite, 0.45, 2, 2, False, False

    If mN = 0 Then
        DRect mL, mT, mPW, mPH, -1, 0, cWhite, 0.5, 1.2 * u
        DText Lp("lum_nolum", "No data"), mL + mPW / 2, mT + mPH / 2, 19 * u, False, cWhite, 0.28, 1, 1, False, False
        FinishChart nm
        Exit Sub
    End If

    ' spec values that may widen the range
    Dim specV(1 To 3) As Double, nSpec As Long
    If mHLo Then nSpec = nSpec + 1: specV(nSpec) = mLo
    If mHHi Then nSpec = nSpec + 1: specV(nSpec) = mHi
    If mTypOn And mHTy Then nSpec = nSpec + 1: specV(nSpec) = mTy

    Dim k As Long, bw As Double, start As Double, span As Double, pad As Double, peak As Double
    Dim bn() As Long, bb() As Long, idx As Long, autoB As Boolean
    If Not isRun Then
        span = mMax - mMin
        autoB = (LCase$(mBins) = "auto" Or Not IsNumV(mBins))
        If autoB Then
            k = -Int(-Sqr(mN))
            If k < 5 Then k = 5
            If k > 25 Then k = 25
        Else
            k = CLng(Val(mBins))
            If k < 1 Then k = 1
        End If
        If span <= 0 Then
            If Abs(mMax) * 0.01 > 0.01 Then bw = NiceStep(Abs(mMax) * 0.01) Else bw = NiceStep(0.01)
            start = mMin - bw / 2: k = 1
        ElseIf autoB Then
            bw = NiceStep(span / k)
            start = Int(mMin / bw + 0.000000001) * bw
            k = Int((mMax - start) / bw + 0.000000001) + 1
            If k < 1 Then k = 1
        Else
            bw = span / k: start = mMin
        End If
        ReDim bn(0 To k - 1): ReDim bb(0 To k - 1)
        For i = 1 To mN
            idx = Int((mLs(i) - start) / bw + 0.000000001)
            If idx < 0 Then idx = 0
            If idx > k - 1 Then idx = k - 1
            bn(idx) = bn(idx) + 1
            If mBad(i) Then bb(idx) = bb(idx) + 1
        Next i
        mX0 = start: mX1 = start + k * bw
        If mNormal And mOkSd Then
            If mAvg - 3.2 * mSd < mX0 Then mX0 = mAvg - 3.2 * mSd
            If mAvg + 3.2 * mSd > mX1 Then mX1 = mAvg + 3.2 * mSd
        End If
        If mIncSpec Then
            For i = 1 To nSpec
                If specV(i) < mX0 Then mX0 = specV(i)
                If specV(i) > mX1 Then mX1 = specV(i)
            Next i
        End If
        pad = (mX1 - mX0) * 0.04: mX0 = mX0 - pad: mX1 = mX1 + pad
        peak = 0
        For i = 0 To k - 1
            If bn(i) > peak Then peak = bn(i)
        Next i
        If mNormal And mOkSd Then
            If mN * bw / (mSd * Sqr(2 * 3.14159265358979)) > peak Then peak = mN * bw / (mSd * Sqr(2 * 3.14159265358979))
        End If
        mY0 = 0: mY1 = peak * 1.3
    Else
        mX0 = 0.5: mX1 = 0
        For i = 1 To mN
            If mSeq(i) > mX1 Then mX1 = mSeq(i)
        Next i
        mX1 = mX1 + 0.5
        If mX1 <= mX0 Then mX1 = mX0 + 1
        mY0 = mMin: mY1 = mMax
        If mIncSpec Then
            For i = 1 To nSpec
                If specV(i) < mY0 Then mY0 = specV(i)
                If specV(i) > mY1 Then mY1 = specV(i)
            Next i
        End If
        span = mY1 - mY0
        If span = 0 Then span = IIf(Abs(mY1) * 0.05 > 1, Abs(mY1) * 0.05, 1)
        mY0 = mY0 - span * 0.1: mY1 = mY1 + span * 0.16
    End If

    Dim xs As Double, ys As Double, xt As Variant, yt As Variant, nxt As Long, nyt As Long
    If isRun Then
        xs = NiceStep((mX1 - mX0) / 10): If xs < 1 Then xs = 1
        ys = NiceStep((mY1 - mY0) / 7)
    Else
        xs = NiceStep((mX1 - mX0) / 8)
        ys = NiceStep((mY1 - mY0) / 6): If ys < 1 Then ys = 1
    End If
    xt = TickList(mX0, mX1, xs, nxt)
    yt = TickList(mY0, mY1, ys, nyt)

    ' grid
    For i = 1 To nxt
        If Not isRun Or xt(i) >= 1 Then DLine XP(xt(i)), mT, XP(xt(i)), mT + mPH, cWhite, 0.87, 1 * u, 0
    Next i
    For i = 1 To nyt
        DLine mL, YP(yt(i)), mL + mPW, YP(yt(i)), cWhite, 0.87, 1 * u, 0
    Next i

    ' fail zones
    Dim a As Double, b As Double
    If mHLo Then
        If isRun Then
            a = YP(mLo): If a < mT Then a = mT
            If a < mT + mPH Then DRect mL, a, mPW, mT + mPH - a, cBad, 0.89, -1, 0, 0
        Else
            b = XP(mLo): If b > mL + mPW Then b = mL + mPW
            If b > mL Then DRect mL, mT, b - mL, mPH, cBad, 0.89, -1, 0, 0
        End If
    End If
    If mHHi Then
        If isRun Then
            b = YP(mHi): If b > mT + mPH Then b = mT + mPH
            If b > mT Then DRect mL, mT, mPW, b - mT, cBad, 0.89, -1, 0, 0
        Else
            a = XP(mHi): If a < mL Then a = mL
            If a < mL + mPW Then DRect a, mT, mL + mPW - a, mPH, cBad, 0.89, -1, 0, 0
        End If
    End If

    Dim px() As Double, py() As Double, xa As Double, xb As Double, g As Double, good As Long, v As Double, kk As Double, r As Double
    If Not isRun Then
        For i = 0 To k - 1
            If bn(i) > 0 Then
                xa = XP(start + i * bw): xb = XP(start + (i + 1) * bw)
                g = (xb - xa) * 0.08: If g > 1.5 * u Then g = 1.5 * u
                good = bn(i) - bb(i)
                If bb(i) > 0 Then DRect xa + g, YP(bb(i)), xb - xa - 2 * g, YP(0) - YP(bb(i)), cBad, 0, cWhite, 0.15, 1 * u
                If good > 0 Then DRect xa + g, YP(bn(i)), xb - xa - 2 * g, YP(bb(i)) - YP(bn(i)), C_OK, 0, cWhite, 0.15, 1 * u
                If xb - xa > 15 * u Then DText CStr(bn(i)), (xa + xb) / 2, YP(bn(i)) - 5 * u, 14 * u, True, cWhite, 0, 1, 2, True, False
            End If
        Next i
        If mNormal And mOkSd Then
            ReDim px(1 To 241): ReDim py(1 To 241)
            kk = mN * bw / (mSd * Sqr(2 * 3.14159265358979))
            For i = 0 To 240
                v = mX0 + (mX1 - mX0) * i / 240
                px(i + 1) = XP(v)
                py(i + 1) = YP(kk * Exp(-0.5 * ((v - mAvg) / mSd) ^ 2))
            Next i
            DPoly px, py, 241, RGB(0, 0, 0), 0.4, 4.4 * u
            DPoly px, py, 241, cFit, 0, 2.4 * u
        End If
    Else
        ReDim px(1 To mN): ReDim py(1 To mN)
        For i = 1 To mN
            px(i) = XP(mSeq(i)): py(i) = YP(mLs(i))
        Next i
        DPoly px, py, mN, cWhite, 0.58, 1.4 * u
        If mN > 300 Then
            r = 2.6 * u
        ElseIf mN > 100 Then
            r = 3.4 * u
        Else
            r = 4.6 * u
        End If
        For i = 1 To mN
            If mBad(i) Then
                DOval px(i), py(i), r + 0.8 * u, RGB(255, 59, 48), cWhite, 1.3 * u
            Else
                DOval px(i), py(i), r, C_OK, cWhite, 1.1 * u
            End If
        Next i
    End If

    ' spec / Typ / mean lines with labels
    Dim kV(1 To 4) As Double, kC(1 To 4) As Long, kD(1 To 4) As Long, kT(1 To 4) As String, kP(1 To 4) As Double, nm4 As Long
    If mHLo Then nm4 = nm4 + 1: kV(nm4) = mLo: kC(nm4) = cBad: kD(nm4) = 4: kT(nm4) = Lp("lum_lsl", "LSL") & " " & FmtN(mLo)
    If mHHi Then nm4 = nm4 + 1: kV(nm4) = mHi: kC(nm4) = cBad: kD(nm4) = 4: kT(nm4) = Lp("lum_usl", "USL") & " " & FmtN(mHi)
    If mTypOn And mHTy Then nm4 = nm4 + 1: kV(nm4) = mTy: kC(nm4) = cTyp: kD(nm4) = 2: kT(nm4) = Lp("lum_typ_t", "Typ") & " " & FmtN(mTy)
    If mMeanOn Then nm4 = nm4 + 1: kV(nm4) = mAvg: kC(nm4) = cMean: kD(nm4) = 0: kT(nm4) = Lp("lum_mean_t", "Mean") & " " & FmtL(mAvg)
    For i = 1 To nm4
        If isRun Then kP(i) = YP(kV(i)) Else kP(i) = XP(kV(i))
    Next i
    ' sort: histogram left to right, trend bottom to top
    Dim tD As Double, tL As Long, tS As String
    For i = 1 To nm4 - 1
        For j = i + 1 To nm4
            If (Not isRun And kP(j) < kP(i)) Or (isRun And kP(j) > kP(i)) Then
                tD = kP(i): kP(i) = kP(j): kP(j) = tD
                tD = kV(i): kV(i) = kV(j): kV(j) = tD
                tL = kC(i): kC(i) = kC(j): kC(j) = tL
                tL = kD(i): kD(i) = kD(j): kD(j) = tL
                tS = kT(i): kT(i) = kT(j): kT(j) = tS
            End If
        Next j
    Next i
    Dim lsz As Double, lh As Double, p As Double, t As String, edg As Long, w As Double, bx As Double, by As Double
    Dim rightSide As Boolean, lastY As Double, lv As Long, clash As Boolean
    Dim plv(1 To 4) As Long, px0(1 To 4) As Double, px1(1 To 4) As Double, np As Long
    lsz = 15 * u: lh = 22 * u: lastY = 1E+30
    For i = 1 To nm4
        p = kP(i): t = kT(i): edg = 0
        If Not isRun Then
            If p < mL Then
                p = mL: t = ChrW(&H25C0) & " " & t: edg = -1
            ElseIf p > mL + mPW Then
                p = mL + mPW: t = t & " " & ChrW(&H25B6): edg = 1
            End If
            If edg = 0 Then
                DLine p, mT, p, mT + mPH, RGB(0, 0, 0), 0.45, 4.2 * u, kD(i)
                DLine p, mT, p, mT + mPH, kC(i), 0, 2.4 * u, kD(i)
            End If
            w = TextW(t, lsz) + 12 * u
            rightSide = (edg = 1) Or (edg = 0 And p + 6 * u + w > mL + mPW)
            If rightSide Then
                bx = p - IIf(edg <> 0, 4, 6) * u - w
            Else
                bx = p + IIf(edg <> 0, 4, 6) * u
            End If
            lv = 0
            Do
                clash = False
                For j = 1 To np
                    If plv(j) = lv And bx < px1(j) + 4 * u And bx + w > px0(j) - 4 * u Then clash = True
                Next j
                If clash Then lv = lv + 1
            Loop While clash
            np = np + 1: plv(np) = lv: px0(np) = bx: px1(np) = bx + w
            by = mT + 8 * u + lv * (lh + 4 * u)
        Else
            If p > mT + mPH Then
                p = mT + mPH: t = ChrW(&H25BC) & " " & t: edg = 1
            ElseIf p < mT Then
                p = mT: t = ChrW(&H25B2) & " " & t: edg = -1
            End If
            If edg = 0 Then
                DLine mL, p, mL + mPW, p, RGB(0, 0, 0), 0.45, 4.2 * u, kD(i)
                DLine mL, p, mL + mPW, p, kC(i), 0, 2.4 * u, kD(i)
            End If
            w = TextW(t, lsz) + 12 * u
            If edg = 1 Then
                by = p - lh - 4 * u
            ElseIf edg = -1 Then
                by = p + 4 * u
            Else
                by = p - lh - 3 * u
            End If
            If by + lh > lastY - 3 * u Then by = lastY - lh - 3 * u
            If by < mT + 2 * u Then by = mT + 2 * u
            lastY = by
            bx = mL + mPW - 8 * u - w
        End If
        DRect bx, by, w, lh, RGB(0, 0, 0), 0.22, kC(i), 0, 1.2 * u
        DText t, bx + 6 * u, by + lh / 2, lsz, True, kC(i), 0, 0, 1, True, False
    Next i

    ' frame, ticks, axis titles
    DRect mL, mT, mPW, mPH, -1, 0, cWhite, 0.2, 1.5 * u
    Dim xd As Long, yd As Long
    If isRun Then xd = 0: yd = LumDec(ys) Else xd = LumDec(xs): yd = 0
    For i = 1 To nxt
        If Not isRun Or xt(i) >= 1 Then
            DLine XP(xt(i)), mT + mPH, XP(xt(i)), mT + mPH + 8 * u, cWhite, 0, 1.3 * u, 0
            DText FmtD(xt(i), xd), XP(xt(i)), mT + mPH + 14 * u, 18 * u, False, cTick, 0, 1, 0, True, False
        End If
    Next i
    For i = 1 To nyt
        DLine mL - 8 * u, YP(yt(i)), mL, YP(yt(i)), cWhite, 0, 1.3 * u, 0
        DText FmtD(yt(i), yd), mL - 13 * u, YP(yt(i)), 18 * u, False, cTick, 0, 2, 1, True, False
    Next i
    DText IIf(isRun, Lp("lum_smp", "Sample No."), Lp("lum_lax", "L")), mL + mPW / 2, mH - 22 * u, 21 * u, True, cWhite, 0, 1, 2, False, False
    DText IIf(isRun, Lp("lum_lax", "L"), Lp("lum_cnt", "Count")), 30 * u, mT + mPH / 2, 21 * u, True, cWhite, 0, 1, 1, False, True

    ' statistics panel
    If panel Then
        Dim itK(1 To 14) As String, itV(1 To 14) As String, itB(1 To 14) As Boolean, nIt As Long
        nIt = 0
        nIt = nIt + 1: itK(nIt) = Lp("lum_n", "n"): itV(nIt) = CStr(mN)
        nIt = nIt + 1: itK(nIt) = Lp("lum_mean_t", "Mean"): itV(nIt) = FmtL(mAvg)
        nIt = nIt + 1: itK(nIt) = Lp("lum_max", "Max"): itV(nIt) = FmtL(mMax)
        nIt = nIt + 1: itK(nIt) = Lp("lum_min", "Min"): itV(nIt) = FmtL(mMin)
        nIt = nIt + 1: itK(nIt) = Lp("lum_sd", "Std dev"): itV(nIt) = IIf(mN > 1, FmtL(mSd), ChrW(&H2014))
        If mHLo Then nIt = nIt + 1: itK(nIt) = Lp("lum_lsl", "LSL"): itV(nIt) = FmtN(mLo)
        If mHHi Then nIt = nIt + 1: itK(nIt) = Lp("lum_usl", "USL"): itV(nIt) = FmtN(mHi)
        If mHTy Then nIt = nIt + 1: itK(nIt) = Lp("lum_typ_t", "Typ"): itV(nIt) = FmtN(mTy)
        If mHasCpk Then nIt = nIt + 1: itK(nIt) = CpkName(): itV(nIt) = Format$(mCpk, "0.00"): itB(nIt) = (mCpk < 1.33)
        If mHLo Then nIt = nIt + 1: itK(nIt) = Lp("lum_below", "Below LSL"): itV(nIt) = CStr(mBelow): itB(nIt) = (mBelow > 0)
        If mHHi Then nIt = nIt + 1: itK(nIt) = Lp("lum_above", "Above USL"): itV(nIt) = CStr(mAbove): itB(nIt) = (mAbove > 0)
        If mHLo Or mHHi Then
            nIt = nIt + 1: itK(nIt) = Lp("lum_yield", "Yield")
            itV(nIt) = Format$((1 - (mBelow + mAbove) / mN) * 100, "0.0") & "%": itB(nIt) = (mBelow + mAbove > 0)
        End If
        Dim lgT(1 To 3) As Long, lgC(1 To 3) As Long, lgS(1 To 3) As String, nLg As Long
        nLg = 1: lgT(1) = 0: lgC(1) = C_OK: lgS(1) = Lp("lum_pass", "Pass")
        If mHLo Or mHHi Then nLg = nLg + 1: lgT(nLg) = 0: lgC(nLg) = cBad: lgS(nLg) = Lp("lum_fail", "Fail")
        If Not isRun And mNormal And mOkSd Then nLg = nLg + 1: lgT(nLg) = 1: lgC(nLg) = cFit: lgS(nLg) = Lp("lum_normal_t", "Normal fit")

        Dim ih As Double, pbx As Double, pbw As Double, head As Double, bh As Double, cy As Double, ly As Double, sx As Double
        ih = 25 * u: pbx = mL + mPW + 18 * u: pbw = mW - pbx - 14 * u: head = 34 * u
        bh = head + nIt * ih + 12 * u + 10 * u + nLg * ih + 8 * u
        DRect pbx, mT, pbw, bh, cWhite, 0.94, cWhite, 0.65, 1 * u
        DText Lp("lum_stats", "Statistics"), pbx + 12 * u, mT + 10 * u, 17 * u, True, cWhite, 0, 0, 0, False, False
        For i = 1 To nIt
            cy = mT + head + ih * (i - 0.5)
            DText itK(i), pbx + 12 * u, cy, 14.5 * u, False, cWhite, 0.26, 0, 1, False, False
            DText itV(i), pbx + pbw - 12 * u, cy, 15.5 * u, True, IIf(itB(i), RGB(255, 107, 97), cWhite), 0, 2, 1, True, False
        Next i
        ly = mT + head + nIt * ih + 12 * u
        DLine pbx + 10 * u, ly, pbx + pbw - 10 * u, ly, cWhite, 0.78, 1 * u, 0
        For i = 1 To nLg
            cy = ly + 10 * u + ih * (i - 0.5): sx = pbx + 22 * u
            If lgT(i) = 0 Then
                DRect sx - 9 * u, cy - 7 * u, 18 * u, 14 * u, lgC(i), 0, cWhite, 0.15, 1 * u
            Else
                DLine sx - 11 * u, cy, sx + 11 * u, cy, lgC(i), 0, 2.6 * u, 0
            End If
            DText lgS(i), pbx + 40 * u, cy, 14.5 * u, False, cWhite, 0, 0, 1, False, False
        Next i
    End If
    FinishChart nm
End Sub

' group everything drawn for one chart and give it its fixed name
Private Sub FinishChart(ByVal nm As String)
    Dim arr() As Variant, j As Long, g As Shape
    If mCnt = 0 Then Exit Sub
    If mCnt = 1 Then
        mSh.Shapes(mNm(1)).Name = nm
    Else
        ReDim arr(1 To mCnt)
        For j = 1 To mCnt: arr(j) = mNm(j): Next j
        Set g = Nothing
        On Error Resume Next
        Set g = mSh.Shapes.Range(arr).Group
        On Error GoTo 0
        If Not g Is Nothing Then g.Name = nm
    End If
    On Error Resume Next
    mSh.Shapes(nm).Placement = 2   ' xlMove: moves with the cells like the exported picture
    On Error GoTo 0
    mCnt = 0
End Sub
