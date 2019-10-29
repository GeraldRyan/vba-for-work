Attribute VB_Name = "Quitter"
Sub CloseAll()

'' closes all workbooks and quits excel. Does not save any that are read only and does not prompt. _
Saves all others.

Application.EnableEvents = False
Application.DisplayAlerts = False
Dim remaining As Long
Dim wb As Workbook

For Each wb In Application.Workbooks
    remaining = Application.Workbooks.count
    If wb.ReadOnly = True Then
        wb.Close savechanges:=False
    Else:
        wb.Close savechanges:=True
    End If
Next

If remaining = 1 Then
   Application.EnableEvents = True
   Application.Quit
End If

End Sub

Sub CloseWithoutSaving()

'' Closes the current active workbook. If read-only, does not prompt to save. If last open, quits application.

Application.EnableEvents = False
Dim bool As Integer
Dim PreRunCount As Integer
Dim window As window

PreRunCount = 0
For Each window In Application.Windows
    If window.Visible = True Then
        PreRunCount = PreRunCount + 1
    End If
Next

On Error Resume Next
    If Application.ActiveWorkbook.ReadOnly = True Then
        Application.ActiveWorkbook.Close savechanges:=False
    Else
        If Application.ActiveWorkbook.Saved = False Then
            bool = MsgBox("This is not read-only and not saved. Are you sure you want to close without saving?", vbYesNo)
            If bool = 6 Then
                Application.ActiveWorkbook.Close savechanges:=False
            End If
        End If
        Application.ActiveWorkbook.Close
    End If

If PreRunCount = 1 Then
    Application.Quit
End If

End Sub

Sub CloseAllReadOnlyWBs() '' Should work but untested.

Application.EnableEvents = False
Dim iOpen As Integer
Dim wb As Workbook

iOpen = Application.Workbooks.count
For Each wb In Application.Workbooks
iOpen = Application.Workbooks.count
    If wb.ReadOnly = True Then
        wb.Close savechanges:=False
        If iOpen = 1 Then
            Application.EnableEvents = True
            Application.Quit
        End If
    End If
Next

Application.EnableEvents = True
End Sub

