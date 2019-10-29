Attribute VB_Name = "RandomSelector"
Sub RandomSample()

UserInput

End Sub

Sub UserInput()

    Dim MyArray As Variant
    Dim i As Integer
    Dim r As Integer
    Dim rng As Range
    Dim RowCount As Integer
    Dim Quantity As Integer
    Set rng = Application.InputBox("Select range of sample", "Select Range", Type:=8)
    Quantity = Application.InputBox("What is your sample size", Type:=1)
    RowCount = rng.rows.count
    'r = RndBetween(1, count)
    ReDim MyArray(Quantity - 1)
    
    For i = 0 To Quantity - 1
        MyArray(i) = RndBetween(1, RowCount)
    Next
    
    MsgBox ("the row count is " & RowCount & " and the sample size is " & Quantity)

End Sub


Function RndBetween(Low, High) As Integer
   Randomize
   RndBetween = Int((High - Low + 1) * Rnd + Low)
End Function
