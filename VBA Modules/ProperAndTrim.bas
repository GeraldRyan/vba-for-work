Attribute VB_Name = "ProperAndTrim"
Sub propercase()

Dim Jimmy As Range, cell As Range

Set Jimmy = Selection

For Each cell In Jimmy

cell.Value = WorksheetFunction.Proper(cell.Value)

Next cell

End Sub

Sub Trimmer()

Dim Rng As Range
Set Rng = Selection
For Each cell In Rng
    cell.Value = WorksheetFunction.Trim(cell)
Next cell

End Sub
