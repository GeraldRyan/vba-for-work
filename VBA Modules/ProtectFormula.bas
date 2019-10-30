Attribute VB_Name = "ProtectFormula"



Sub lockCellsWithFormulas()
With ActiveSheet
    .Unprotect
    .Cells.Locked = False
    .Cells.SpecialCells(xlCellTypeFormulas).Locked = True
    .Protect AllowDeletingRows:=True
End With
End Sub

Sub highlightMaxValue()
Dim rng As range
For Each rng In Selection
    If rng = WorksheetFunction.Max(Selection) Then
        rng.Style = "Good"
    End If
Next rng
End Sub

Sub highlightMinValue()

Dim rng As range

For Each rng In Selection
    If rng = WorksheetFunction.Min(Selection) Then
    rng.Style = "Good"
    End If
Next rng

End Sub

