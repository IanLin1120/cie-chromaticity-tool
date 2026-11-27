' Luminance sheet (A19 = No.): changing an L value (column C) or the spec row (B6:D6) redraws its charts.
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Not Module2.IsLumSheet(Sh) Then Exit Sub
    If Application.Intersect(Target, Sh.Range("B6:D6,A20:D" & Sh.Rows.Count)) Is Nothing Then Exit Sub
    Module2.LumAuto Sh
End Sub
