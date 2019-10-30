Attribute VB_Name = "AnnettesMedicalMacro"
Option Explicit

Sub main()

Application.EnableEvents = False
Application.ScreenUpdating = False

'declare variables
Dim PivotTableWS As Worksheet
Set PivotTableWS = Application.ActiveSheet
Dim ws As Worksheet
Dim looptimes As Integer
Dim cell As range
Dim ptlname As String


For Each ws In Application.ActiveWorkbook.Worksheets
    If ws.name = "Summary" Or ws.name = PivotTableWS.name Or ws.Visible = False Then
        GoTo next1
    End If
    
    ' initialize value to zero
    For Each cell In ws.range("B:B")
        If IsNumeric(cell.Value) And cell.Value <> "" And InStr(cell.FormulaR1C1, "=") = 0 Then
            cell.FormulaR1C1 = 0
        End If
    Next
    
    For Each cell In PivotTableWS.UsedRange
    
        'check if a real name but not their totals
        If InStr(cell.Value, ",") <> 0 And cell.Offset(0, 1).Value <> "" Then
            ptlname = Mid(cell.Value, InStr(cell.Value, ",") + 2, 1) & ". " & Left(cell.Value, InStr(cell.Value, ",") - 1)
            
            ' should call once for each WS to populate, 15-20 times PER CELL IN ACTIVE RANGE. TOO MUCH CALLING
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
End Sub


Sub TransferData(PivotTableWS As Worksheet, ws As Worksheet, name As range)
Dim looptimes As Integer
Dim datatotransfer As Long
Dim i As Integer
Dim indexnumber As Long
Dim CellToPasteTo As range


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


Function loopvalues(PivotTableWS As Worksheet, ws As Worksheet, cell As range)
    Dim endrow As Integer
    Dim startrow As Integer
    startrow = cell.row
    endrow = cell.End(xlDown).row
    loopvalues = endrow - startrow

End Function


Sub ZeroOut(ws As Worksheet)
Application.ScreenUpdating = False
Application.EnableEvents = False

Dim PivotTableWS As Worksheet
Dim cell As range


If Not ws.name = "Summary" Or ws.name = PivotTableWS.name Or ws.Visible = False Then
    GoTo endd
End If

For Each cell In ws.range("B:B")
    If IsNumeric(cell.Value) And cell.Value <> "" And InStr(cell.FormulaR1C1, "=") = 0 Then
        cell.FormulaR1C1 = 0
    End If
Next


Application.ScreenUpdating = True
Application.EnableEvents = True

endd:
End Sub
