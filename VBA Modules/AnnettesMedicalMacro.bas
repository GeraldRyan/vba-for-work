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
Dim cell As Range
Dim ptlname As String


For Each ws In Application.ActiveWorkbook.Worksheets
    If ws.name = "Summary" Or ws.name = PivotTableWS.name Or ws.Visible = False Then
    GoTo Next1
    End If
    
    ' initialize value to zero
    For Each cell In ws.Range("B:B")
        If IsNumeric(cell.Value) And cell.Value <> "" And InStr(cell.FormulaR1C1, "=") = 0 Then
            cell.FormulaR1C1 = 0
        End If
    Next
    
    For Each cell In PivotTableWS.UsedRange
        If InStr(cell.Value, ",") <> 0 And cell.Offset(0, 1).Value <> "" Then
            ptlname = Mid(cell.Value, InStr(cell.Value, ",") + 2, 1) & ". " & Left(cell.Value, InStr(cell.Value, ",") - 1)
            If ptlname = ws.name Then
                Call TransferData(PivotTableWS, ws, cell)
            End If
        End If
    Next
Next1:
Next
Application.EnableEvents = True
Application.ScreenUpdating = True

MsgBox ("Have a nice day")
End Sub


Sub TransferData(PivotTableWS As Worksheet, ws As Worksheet, name As Range)
Dim looptimes As Integer
Dim datatotransfer As Long
Dim i As Integer
Dim indexnumber As Long
Dim celltopasteto As Range


    looptimes = loopvalues(PivotTableWS, ws, name)
    'MsgBox ("The name is : " & name.Value & looptimes) 'and the code and hours are " & cell.Offset(0, 1).Value & " and " & cell.Offset(0, 2).Value)

    ' initialize all preexisting data to zero
    



    ' Transfer data looptimes times
    For i = 0 To looptimes - 1
        indexnumber = name.Offset(i, 1).FormulaR1C1
        datatotransfer = name.Offset(i, 2).FormulaR1C1
        Set celltopasteto = ws.UsedRange.Find(indexnumber).Offset(0, 1)
        celltopasteto.FormulaR1C1 = datatotransfer
               
    Next
End Sub


Function loopvalues(PivotTableWS As Worksheet, ws As Worksheet, cell As Range)
    Dim endrow As Integer
    Dim startrow As Integer
    startrow = cell.Row
    endrow = cell.End(xlDown).Row
    loopvalues = endrow - startrow

End Function
