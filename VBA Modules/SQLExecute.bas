Attribute VB_Name = "SQLExecute"
Sub SQLExecute()

    '' This code was from online at AnalystCave.com using SQL in excel. Should be fun to learn and _
    practice the syntax and formatting/usage of SQL in excel

    'Attribute ExecuteSQL.VB_ProcData.VB_Invoke_Func = "S\n14"  '' supposedly a shortcut key
    'AnalystCave.com
    On Error GoTo ErrorHandl
    Dim SQL As String, sConn As String, qt As QueryTable
    SQL = InputBox("Provide your SQL Query", "Run SQL Query")
    If SQL = vbNullString Then Exit Sub
    sConn = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;;Password=;User ID=Admin;Data Source=" & _
        ThisWorkbook.path & "/" & ThisWorkbook.name & ";" & _
        "Mode=Share Deny Write;Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
    Set qt = ActiveCell.Worksheet.QueryTables.add(Connection:=sConn, Destination:=ActiveCell)
    With qt
        .CommandType = xlCmdSql
        .CommandText = SQL
        .name = Int((1000000000 - 1 + 1) * Rnd + 1)
        .RefreshStyle = xlOverwriteCells
        .Refresh BackgroundQuery:=False
    End With
    Exit Sub
ErrorHandl: MsgBox "Error: " & Err.Description: Err.Clear
End Sub

