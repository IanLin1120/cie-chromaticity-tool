' Data sheet events: K2:M2 work like check boxes, J2 picks the diagram image,
' and changing the spec limits (B6:C7) redraws the spec box.
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Row <> 2 Then Exit Sub
    If Target.Rows.Count > 1 Or Target.Columns.Count > 1 Then Exit Sub
    If Target.Column < 11 Or Target.Column > 13 Then Exit Sub
    Application.EnableEvents = False
    Module1.ToggleBox Target
    Me.Range("I2").Select
    Application.EnableEvents = True
    Module1.RefreshChart
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("J2,B6:C7")) Is Nothing Then Exit Sub
    Module1.RefreshChart
End Sub
