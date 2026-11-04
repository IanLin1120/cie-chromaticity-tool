' Data sheet events: K2:O2 work like check boxes, J2 picks the diagram image (both only show/hide
' objects that already exist, so they are instant); changing the spec limits (B6:C7) rebuilds the overlays.
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Row <> 2 Then Exit Sub
    If Target.Rows.Count > 1 Or Target.Columns.Count > 1 Then Exit Sub
    If Target.Column < 11 Or Target.Column > 15 Then Exit Sub
    Application.EnableEvents = False
    Module1.ToggleBox Target
    Me.Range("I2").Select
    Application.EnableEvents = True
    Module1.ApplyView
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("B6:C7")) Is Nothing Then
        Module1.RefreshChart
    ElseIf Not Intersect(Target, Me.Range("J2")) Is Nothing Then
        Module1.ApplyView
    End If
End Sub
