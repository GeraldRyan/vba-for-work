Attribute VB_Name = "learningmodule"
'Sub FileBackUp()

'ThisWorkbook.SaveCopyAs Filename:=ThisWorkbook.Path & "" & Format(Date, "mm-dd-yy") & "" & ThisWorkbook.Name

'End Sub

 'Sub Workbook_BeforeClose(Cancel As Boolean)
  '  Dim sFileName As String
   ' Dim sDateTime As String
'
 '   With ThisWorkbook
  '      sDateTime = " (" & Format(Now, "yyyy-mm-dd hhmm") & ").xlsm"
   '     sFileName = Application.WorksheetFunction.Substitute _
    '      (.FullName, ".xlsm", sDateTime)
     '   .SaveCopyAs sFileName
   ' End With
'End Sub


'Private Sub Workbook_Open()
 '   MsgBox "Welcome"
'End Sub

Sub HideWorksheets()
Attribute HideWorksheets.VB_ProcData.VB_Invoke_Func = "H\n14"
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
If ws.name <> ActiveWorkbook.ActiveSheet.name Then
ws.Visible = xlSheetHidden
End If

Next ws

End Sub



    
Sub UnhideAllWorksheet()
Attribute UnhideAllWorksheet.VB_ProcData.VB_Invoke_Func = "U\n14"

Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Visible = xlSheetVisible
Next ws
End Sub
    

Sub ProtecAllWorskeets()
Dim ws As Worksheet
Dim ps As String
ps = InputBox("Enter a Password.", vbOKCancel)
For Each ws In ActiveWorkbook.Worksheets
ws.Protect Password:=ps
Next ws
End Sub

Sub RemoveSpaces()
Dim myRange As Range
Dim myCell As Range
Select Case MsgBox("You Can't Undo This Action. " & "Save Workbook First?", _
vbYesNoCancel, "Alert")
Case Is = vbYesThisWorkbook.Save
Case Is = vbCancel
Exit Sub
End Select
Set myRange = Selection
For Each myCell In myRange
If Not IsEmpty(myCell) Then
myCell = Trim(myCell)
End If
Next myCell
End Sub

Sub FirstA()
Attribute FirstA.VB_ProcData.VB_Invoke_Func = "P\n14"
Dim sht As Worksheet, csheet As Worksheet


Set csheet = ActiveSheet

For Each sht In ActiveWorkbook.Worksheets
  If sht.Visible Then
    sht.Activate
    Range("A1").Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
  End If
Next sht

csheet.Activate
Worksheets(1).Activate

End Sub

Sub InsertMultipleSheets()
Dim i As Integer
i = InputBox("Enter number of sheets to insert.", "Enter Multiple Sheets")
Sheets.add after:=ActiveSheet, count:=i
End Sub
Sub OpenxCalculator()
Attribute OpenxCalculator.VB_ProcData.VB_Invoke_Func = "C\n14"
Application.ActivateMicrosoftApp index:=0
End Sub

Sub OpenWorkbookAsAttachment()
Attribute OpenWorkbookAsAttachment.VB_ProcData.VB_Invoke_Func = "Z\n14"
Application.Dialogs(xlDialogSendMail).Show
End Sub


Sub highlightNegativeNumbers()
Dim rng As Range
For Each rng In Selection
    If WorksheetFunction.IsNumber(rng) Then
        If rng.Value < 0 Then
            rng.Font.Color = -16776961
        End If
    End If
    Next
End Sub



Sub printSelection()
Selection.PrintOut Copies:=1, Collate:=True
End Sub
