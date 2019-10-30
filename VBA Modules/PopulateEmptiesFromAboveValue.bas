Attribute VB_Name = "PopulateEmptiesFromAboveValue"
Sub PopulateEmptyCellsWithValueDirectlyAbove()
Dim rng As range
Set rng = Selection


For Each cell In rng
    If cell = "" And cell.Offset(0, 1) <> "" Then
    '' copy and paste the value above
        cell.Offset(-1, 0).Select
        Selection.Copy
        Selection.Offset(1, 0).Select
        Application.ActiveSheet.Paste
    End If
Next cell



End Sub

