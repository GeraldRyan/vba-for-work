Attribute VB_Name = "A1All"
Sub A1AllWorksheets()
'Update 20140317
Dim ws As Worksheet
'Application.ScreenUpdating = False
For Each ws In Application.ActiveWorkbook.Worksheets
    ws.Activate
    With Application.ActiveWindow
    ws.[a1].Select
    End With
Next
ActiveWorkbook.Worksheets(1).Activate
'Application.ScreenUpdating = True
End Sub
