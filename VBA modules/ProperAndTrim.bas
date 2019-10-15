Attribute VB_Name = "ProperAndTrim"
Sub propercase()

Dim Jimmy As Range, cell As Range

Set Jimmy = Selection

For Each cell In Jimmy

cell.Value = WorksheetFunction.Proper(cell.Value)

Next cell

End Sub

Sub Trimmer()

Dim rng As Range
Set rng = Selection
For Each cell In rng
    cell.Value = WorksheetFunction.Trim(cell)
Next cell

End Sub
