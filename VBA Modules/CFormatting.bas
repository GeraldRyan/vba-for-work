Attribute VB_Name = "CFormatting"
Sub BlankYellowEach()
Dim cell As range
    For Each cell In Selection
    cell.FormatConditions.add Type:=xlExpression, Formula1:= _
        "=LEN(TRIM(" & cell.Address & "))=0"
    cell.FormatConditions(cell.FormatConditions.count).SetFirstPriority
    With cell.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
    End With
    cell.FormatConditions(1).StopIfTrue = False
    Next cell

End Sub

Sub RedGreenOrange()

'' TODO make color selection dynamic, user focused (color and value selection)

    Selection.FormatConditions.add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=-5"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=-5", Formula2:="=0"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
End Sub

