Attribute VB_Name = "CalculateRunTimeForAnnette"
Option Explicit

Sub CalculateRunTime_Seconds()
'PURPOSE: Determine how many seconds it took for code to completely run
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim StartTime As Double
Dim SecondsElapsed As Double

'Remember time when macro starts
  StartTime = Timer


Application.EnableEvents = False
Application.ScreenUpdating = False

'declare variables
Dim PivotTableWS As Worksheet
Set PivotTableWS = Application.ActiveSheet
Dim ws As Worksheet
Dim looptimes As Integer
Dim cell As Range
Dim ptlname As String

' initialize value to zero (This loop runs about 15 times)
For Each ws In Application.ActiveWorkbook.Worksheets
    If ws.name = "Summary" Or ws.name = PivotTableWS.name Or ws.Visible = False Then GoTo next1
    Call ZeroOut(ws)


    For Each cell In PivotTableWS.UsedRange
    
        'check if a real name but not their totals (c. 50 instances)
        If InStr(cell.Value, ",") <> 0 And cell.Offset(0, 1).Value <> "" Then
            
            ' record the name (50*15)
            ptlname = Mid(cell.Value, InStr(cell.Value, ",") + 2, 1) & ". " & Left(cell.Value, InStr(cell.Value, ",") - 1)
            
            ' should call once for each WS to populate, 15-20 times per real names in pivot table (c.50)
            If ptlname = ws.name Then
                Call TransferData(PivotTableWS, ws, cell)
            End If
        End If
    Next


next1:
Next
    

Application.EnableEvents = True
Application.ScreenUpdating = True

MsgBox ("Have a nice day")
'Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)

'Notify user in seconds
  MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
End Sub


Sub TransferData(PivotTableWS As Worksheet, ws As Worksheet, name As Range)
Dim looptimes As Integer
Dim datatotransfer As Long
Dim i As Integer
Dim indexnumber As Long
Dim CellToPasteTo As Range


    looptimes = loopvalues(PivotTableWS, ws, name)
    'MsgBox ("The name is : " & name.Value & looptimes) 'and the code and hours are " & cell.Offset(0, 1).Value & " and " & cell.Offset(0, 2).Value)

    ' initialize all preexisting data to zero
    



    ' Transfer data looptimes times
    For i = 0 To looptimes - 1
        indexnumber = name.Offset(i, 1).FormulaR1C1
        datatotransfer = name.Offset(i, 2).FormulaR1C1
        Set CellToPasteTo = ws.UsedRange.Find(indexnumber).Offset(0, 1)
        CellToPasteTo.FormulaR1C1 = datatotransfer
               
    Next
End Sub


Function loopvalues(PivotTableWS As Worksheet, ws As Worksheet, cell As Range)
    Dim endrow As Integer
    Dim startrow As Integer
    startrow = cell.row
    endrow = cell.End(xlDown).row
    loopvalues = endrow - startrow

End Function


Sub ZeroOut(ws As Worksheet)

Dim cell As Range
Dim LastRow As Long
Dim rng As Range
LastRow = ws.UsedRange.rows.count
Set rng = ws.Range("B1:B" & LastRow)

If ws.name = "Summary" Or ws.Visible = False Then
    GoTo endd
End If

For Each cell In rng
    If IsNumeric(cell.Value) And cell.Value <> "" And InStr(cell.FormulaR1C1, "=") = 0 Then
        cell.FormulaR1C1 = 0
    End If
Next



endd:
End Sub
