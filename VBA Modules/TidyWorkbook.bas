Attribute VB_Name = "TidyWorkbook"

Sub UnFreeze()
'Update 20140317
Dim ws As Worksheet
Application.ScreenUpdating = False
For Each ws In Application.ActiveWorkbook.Worksheets
    ws.Activate
    With Application.ActiveWindow
            .FreezePanes = False
            .Zoom = 100
            .DisplayGridlines = False
    
    ws.Cells.Font.Size = "8"
    ws.[a1].Select
    ActiveWorkbook.Worksheets(1).Activate
    End With
Next
Application.ScreenUpdating = True
End Sub

Sub TogglePageBreaks()

    If Application.ActiveSheet.DisplayPageBreaks = False Then
        Application.ActiveSheet.DisplayPageBreaks = True
    Else
        Application.ActiveSheet.DisplayPageBreaks = False
    End If
    
End Sub
