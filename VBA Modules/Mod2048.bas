Attribute VB_Name = "Mod2048"
Option Explicit

Const miSIZE As Integer = 4
Dim mrBoard As Range
Dim mrScore As Range
Dim mrHighscore As Range
Public mrMessage As Range

Public bInitialized As Boolean

Public Const iRowKeys = 16
Public Const iColKeys = 9

Public Const gstrPAUSE = "Audit"
Public Const gstrRUNNING = "Going concern"
Public Const gstrENDGAME = "End of FY"
Public Const gstrHIGHSCORE = "Highest profit"

Sub Initialize()
    Set mrBoard = Range("Board")
    Set mrScore = Range("Score")
    Set mrHighscore = Range("Highscore")
    Set mrMessage = Range("Message")
    mrMessage = gstrRUNNING
    isEndgame
    bInitialized = True
    RestorePosition
End Sub

Sub ResetGame()
    Initialize
    mrMessage = gstrRUNNING
    mrBoard.Value = ""
    
    If Not IsNumeric(mrHighscore.Value) Then mrHighscore.Value = 0
    If Not IsNumeric(mrScore.Value) Then mrScore.Value = 0
    
    If mrHighscore.Value < mrScore.Value Then mrHighscore = mrScore
    mrScore = 0
    
    AddTile
    AddTile
End Sub
Sub AddScore(iValue As Integer)
    mrScore = mrScore + iValue
End Sub
Sub RestorePosition()
    Application.EnableEvents = False
        wksBoard.Cells(iRowKeys, iColKeys).Select
    Application.EnableEvents = True
End Sub

Sub PrepareToFromStep(bLeftUp As Boolean, iFrom As Integer, iTo As Integer, iStep As Integer)
    If bLeftUp Then
        iFrom = 1
        iTo = miSIZE - 1
        iStep = 1
    Else
        iFrom = miSIZE
        iTo = 2
        iStep = -1
    End If
End Sub

Function MoveHorizontally(bLeft As Boolean, iRow As Integer) As Boolean
    Dim iCol As Integer
    Dim iCol2 As Integer
    
    Dim iFrom As Integer
    Dim iTo As Integer
    Dim iStep As Integer
    
    PrepareToFromStep bLeft, iFrom, iTo, iStep
    
    For iCol = iFrom To iTo Step iStep
        If mrBoard.Cells(iRow, iCol) = "" Then
            For iCol2 = iCol + iStep To iTo + iStep Step iStep
                If mrBoard.Cells(iRow, iCol2) <> "" Then
                    mrBoard.Cells(iRow, iCol) = mrBoard.Cells(iRow, iCol2)
                    mrBoard.Cells(iRow, iCol2) = ""
                    MoveHorizontally = True
                    GoTo NextCol
                End If
            Next iCol2
        End If
NextCol:
    Next iCol
End Function

Function MergeHorizontally(bLeft As Boolean, iRow As Integer) As Boolean
    Dim iCol As Integer
    
    Dim iFrom As Integer
    Dim iTo As Integer
    Dim iStep As Integer
    
    PrepareToFromStep bLeft, iFrom, iTo, iStep
    
    For iCol = iFrom To iTo Step iStep
        If mrBoard.Cells(iRow, iCol) <> "" Then
            If mrBoard.Cells(iRow, iCol + iStep) = mrBoard.Cells(iRow, iCol) Then
                mrBoard.Cells(iRow, iCol) = 2 * mrBoard.Cells(iRow, iCol)
                AddScore mrBoard.Cells(iRow, iCol)
                mrBoard.Cells(iRow, iCol + iStep) = ""
                MergeHorizontally = True
            End If
        End If
    Next iCol
End Function

Function MoveVertically(bUp As Boolean, iCol As Integer) As Boolean
    Dim iRow As Integer
    Dim iRow2 As Integer
    
    Dim iFrom As Integer
    Dim iTo As Integer
    Dim iStep As Integer
    
    PrepareToFromStep bUp, iFrom, iTo, iStep
    
    For iRow = iFrom To iTo Step iStep
        If mrBoard.Cells(iRow, iCol) = "" Then
            For iRow2 = iRow + iStep To iTo + iStep Step iStep
                If mrBoard.Cells(iRow2, iCol) <> "" Then
                    mrBoard.Cells(iRow, iCol) = mrBoard.Cells(iRow2, iCol)
                    mrBoard.Cells(iRow2, iCol) = ""
                    MoveVertically = True
                    GoTo NextRow
                End If
            Next iRow2
        End If
NextRow:
    Next iRow
End Function

Function MergeVertically(bUp As Boolean, iCol As Integer) As Boolean
    Dim iRow As Integer
    
    Dim iFrom As Integer
    Dim iTo As Integer
    Dim iStep As Integer
    
    PrepareToFromStep bUp, iFrom, iTo, iStep
    
    For iRow = iFrom To iTo Step iStep
        If mrBoard.Cells(iRow, iCol) <> "" Then
            If mrBoard.Cells(iRow + iStep, iCol) = mrBoard.Cells(iRow, iCol) Then
                mrBoard.Cells(iRow, iCol) = 2 * mrBoard.Cells(iRow, iCol)
                AddScore mrBoard.Cells(iRow, iCol)
                mrBoard.Cells(iRow + iStep, iCol) = ""
                MergeVertically = True
            End If
        End If

    Next iRow
End Function

Sub LeftClick()
    Horizontal True
End Sub
Sub RightClick()
    Horizontal False
End Sub
Sub UpClick()
    Vertical True
End Sub
Sub DownClick()
    Vertical False
End Sub

Sub Horizontal(bLeft As Boolean)
    Dim iRow As Integer
    Dim bAddTile As Boolean
    For iRow = 1 To miSIZE
        bAddTile = bAddTile Or (MoveHorizontally(bLeft, iRow))
        bAddTile = bAddTile Or (MergeHorizontally(bLeft, iRow))
        bAddTile = bAddTile Or (MoveHorizontally(bLeft, iRow))
    Next iRow
    If bAddTile Then
        AddTile
        isEndgame
    End If
End Sub

Sub isEndgame()
    If getNoOfEmptyTiles > 0 Then Exit Sub
    
    Dim iCol As Integer
    Dim iRow As Integer
    For iCol = 1 To miSIZE
        For iRow = 1 To miSIZE - 1
            If mrBoard.Cells(iRow, iCol) = mrBoard.Cells(iRow + 1, iCol) Then
                Exit Sub
            End If
        Next iRow
    Next iCol
    
    For iRow = 1 To miSIZE
        For iCol = 1 To miSIZE - 1
            If mrBoard.Cells(iRow, iCol) = mrBoard.Cells(iRow, iCol + 1) Then
                Exit Sub
            End If
        Next iCol
    Next iRow
    
    bInitialized = False
    
    If mrScore.Value > mrHighscore.Value Then
        mrMessage = gstrHIGHSCORE
    Else
        mrMessage = gstrENDGAME
    End If
   
End Sub

Sub Vertical(bUp As Boolean)
     
    Dim iCol As Integer
    Dim bAddTile As Boolean
    For iCol = 1 To miSIZE
        bAddTile = bAddTile Or (MoveVertically(bUp, iCol))
        bAddTile = bAddTile Or (MergeVertically(bUp, iCol))
        bAddTile = bAddTile Or (MoveVertically(bUp, iCol))
    Next iCol
    If bAddTile Then
        AddTile
        isEndgame
    End If
End Sub
Function getNoOfEmptyTiles() As Integer
    getNoOfEmptyTiles = miSIZE * miSIZE - WorksheetFunction.CountA(mrBoard)
End Function
Sub AddTile()
    Dim iEmptySpaces As Integer
    iEmptySpaces = getNoOfEmptyTiles
    If iEmptySpaces = 0 Then
        
        Exit Sub
    Else
        
        
        Dim iNo As Integer
        iNo = WorksheetFunction.RandBetween(1, iEmptySpaces)
        
        Dim rTile As Range
        For Each rTile In mrBoard
            If rTile.Value = "" Then
                iNo = iNo - 1
                If iNo = 0 Then
                    rTile.Value = WorksheetFunction.RandBetween(1, 2) * 2
                    Exit Sub
                End If
            End If
        Next rTile
        
    End If
End Sub

