Attribute VB_Name = "ColumnWidthAutoFit"
Sub ColumnWidthAutoFit()
Dim rng As Range

'autofit the columns of the selection
Selection.EntireColumn.AutoFit

' Make the empty columns the default width
For Each rng In Selection.Columns
    If Application.CountA(rng.EntireColumn) = 0 Then
        rng.EntireColumn.ColumnWidth = 2 '8.43 (default)
    End If
Next

End Sub

Sub CallColumnWidthAutoFit(rng As Range)

'autofit the columns of the selection
Selection.EntireColumn.AutoFit

' Make the empty columns the default width
For Each rng In Selection.Columns
    If Application.CountA(rng.EntireColumn) = 0 Then
        rng.EntireColumn.ColumnWidth = 2  '8.43 (default)
    End If
Next

End Sub

