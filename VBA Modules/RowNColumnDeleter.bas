Attribute VB_Name = "RowNColumnDeleter"
Sub RowNColumnDeleter()
'
' DeleteRow Macro
'
' Deletes rows and columns in a selection.
' Warning, Be careful not to delete good data
' If One column is selected, deletes all the rows in selection
' Otherwise it deletes all the columns in the selection
'


Dim MySelection As Range
Dim RowsToDelete As Range
Dim ColumnsToDelete As Range

    Set MySelection = Selection
    If MySelection.Columns.count = 1 Then
        Set RowsToDelete = MySelection.rows.EntireRow
        RowsToDelete.Delete
    Else
        Set ColumnsToDelete = MySelection.Columns.EntireColumn
        ColumnsToDelete.Delete
    End If
    
    
End Sub
 

Sub DeleteJustRows()
'
' DeleteRow Macro
'
' Deletes rows in a selection.
' Warning, Be careful not to delete good data as there is no undo button

Application.EnableEvents = False
Selection.rows.EntireRow.Delete
Application.EnableEvents = True


End Sub
