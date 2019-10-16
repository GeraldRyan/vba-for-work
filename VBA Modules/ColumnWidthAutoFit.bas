Attribute VB_Name = "ColumnWidthAutoFit"
Sub ColumnWidthAutoFit()
Dim rng As Range

'autofit the columns of the selection
Selection.EntireColumn.AutoFit

' Make the empty columns the default width
For Each rng In Selection.Columns
    If Application.CountA(rng.EntireColumn) = 0 Then
        rng.EntireColumn.ColumnWidth = 8.43
    End If
Next

End Sub

