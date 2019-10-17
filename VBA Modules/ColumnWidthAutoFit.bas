Attribute VB_Name = "ColumnWidthAutoFit"
Sub ColumnWidthAutoFit()
Dim Rng As Range

'autofit the columns of the selection
Selection.EntireColumn.AutoFit

' Make the empty columns the default width
For Each Rng In Selection.Columns
    If Application.CountA(Rng.EntireColumn) = 0 Then
        Rng.EntireColumn.ColumnWidth = 8.43
    End If
Next

End Sub

