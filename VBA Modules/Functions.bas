Attribute VB_Name = "Functions"
Public Function add(n1 As Double, n2 As Double)
add = n1 + n2
End Function

Public Function Percent(n1 As Double, n2 As Double)
Percent = (n1 - n2) / n2
End Function


Public Function Difference(n1 As Double, n2 As Double)
Difference = (n2 - n1)
End Function


Public Function Average2(n1 As Double, n2 As Double)
Average2 = (n1 + n2) / 2

End Function

Function GetCombination(CoinsRange As range, SumCellId As range) As String
Dim Nb As Integer
Dim Com As String
Dim Sum As Double
Dim r As range
Set r = CoinsRange
Sum = SumCellId.Value
For Each cell In r.Cells
If Sum / cell.Value >= 1 Then
Com = Com & Int(Sum / cell.Value) & " of " & cell.Value & "  "
Sum = Sum - (Int(Sum / cell.Value)) * cell.Value
End If
Next
GetCombination = Com
End Function
