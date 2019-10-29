Attribute VB_Name = "ProperAndTrim"
Sub propercase()

Dim Jimmy As Range, Buffett As Range

Set Jimmy = Selection

For Each Buffett In Jimmy

Buffett.Value = WorksheetFunction.Proper(Buffett.Value)

Next Buffett

End Sub

Sub Trimmer()

Dim rng As Range
Set rng = Selection
For Each cell In rng
    cell.Value = WorksheetFunction.Trim(cell)
Next cell

End Sub
