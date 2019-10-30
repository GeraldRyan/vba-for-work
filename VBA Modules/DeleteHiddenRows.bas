Attribute VB_Name = "DeleteHiddenRows"
Sub DeleteHiddenRows()
Dim SelectedArea As range
Set SelectedArea = Selection

For Each rw In SelectedArea.rows

If rows(rw.row).Hidden = True Then rows(rw.row).EntireRow.Delete
Next

End Sub
