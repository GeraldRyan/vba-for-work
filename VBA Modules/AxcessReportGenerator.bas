Attribute VB_Name = "AxcessReportGenerator"
Option Explicit
Sub AxcessTabler()

Dim filePath As String
filePath = "C:\Users\gcr\Documents\Clients\DSB\Elaine\Axcess Reporter\ServiceCodeTable.csv"
Dim cell As Range
Dim codeValue As String

Dim valueToPaste As String

For Each cell In Selection
    '' TODO validate cell is a serviceCode or optimize
    codeValue = cell.Value
    '' reference lookup table
    valueToPaste = ReadFile(filePath, codeValue)
    cell.Offset(0, 1).Value = valueToPaste

Next

End Sub

Function ReadFile(filePath As String, serviceCode As String) As String
Dim lineFromFile As Variant
Dim lineItems As Variant

Open filePath For Input As #1
Do Until EOF(1)
    Line Input #1, lineFromFile
    lineItems = Split(lineFromFile, ",")

    If lineItems(2) = serviceCode Or lineItems(2) = Left(serviceCode, Len(serviceCode) - 1) Then
        ReadFile = lineItems(3)
        GoTo endThis
    End If
Loop

endThis:
Close #1
End Function
