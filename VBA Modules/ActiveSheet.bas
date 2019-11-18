Attribute VB_Name = "ActiveSheet"
Sub ActiveSheetDelete()
'
' deleteWS Macro
'

Application.ActiveSheet.Delete
End Sub

Sub CopySheet()

Application.ActiveSheet.Copy after:=Worksheets(Application.ActiveSheet.index)

End Sub
