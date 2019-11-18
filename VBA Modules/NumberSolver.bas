Attribute VB_Name = "NumberSolver"
Option Explicit
'This is code from online to explore to see if it solves for values.
'Very hard coded not so dynamic, may have to make dynamic.


Public Function DecToBin(ByVal DecimalIn As Variant) As String
        ' The DecimalIn argument is limited to 79228162514264337593543950245
        ' (approximately 96-bits) - large numerical values must be entered
        ' as a String value to prevent conversion to scientific notation.

        'Function Created by Rick Rothstein, of MrExcel.com, 04/2011
        DecToBin = ""
        DecimalIn = CDec(DecimalIn)
        Do While DecimalIn <> 0
                DecToBin = Trim$(Str$(DecimalIn - 2 * Int(DecimalIn / 2))) & DecToBin
                DecimalIn = Int(DecimalIn / 2)
        Loop
End Function

Sub Find_Combination()
        
        'Author: Christian Loveridge
        'christian.loveridge@gmail.com
        
        'Finds the combination of unique numbers which add up to a target value
        
        If Range("D7").Value = "" Then
                MsgBox "You need to add in variables!"
                End
        End If
        
        If Check_Values_Are_Clean Then
        Else
                MsgBox "One or more values are not numbers -- this will break the program. Please remove these and fill in the blanks."
                End
        End If
        
        
        
        Range("Attempt_Number").Value = 1
        
        If MsgBox("This can run for a long time, especially with many variables. Continue?", vbYesNo) = vbNo Then
                End
        End If
        
        Application.ScreenUpdating = False
        
        Range("I4").Value = TimeSerial(Hour(Now), Minute(Now), Second(Now))
        Range("J4").FormulaR1C1 = "=TEXT(NOW(),""hh:mm:ss"")"
        Range("K4").FormulaR1C1 = "=TEXT(RC[-1]-RC[-2],""hh:mm:ss"")"
        
        Dim mx_bit As Long
        mx_bit = Max_Bit
        
        Dim attempt As Long
        attempt = 0
        
        Do While attempt <= mx_bit
                DoEvents
                
                If attempt Mod 1000 = 0 Then
                        Application.ScreenUpdating = True
                        Application.Calculate
                        Application.ScreenUpdating = False
                End If
                
                If Range("Sum_Total").Value <> Range("Target_Value").Value Then
                        attempt = attempt + 1
                        Range("Attempt_Number").Value = "'" & attempt
                Else
                        Range("J4").Value = TimeSerial(Hour(Now), Minute(Now), Second(Now))
                        MsgBox "Found a solution!"
                        
                        End
                End If
        Loop
        
        Range("J4").Value = TimeSerial(Hour(Now), Minute(Now), Second(Now))
        MsgBox "The target value is not made from any combination of these numbers..."
        
End Sub

Private Function Max_Bit() As String
        'finds the maximum # of attempts to make before ending the search
        Dim check_row As Long
        check_row = 7
        Max_Bit = 1
        
        Do While Range("D" & check_row).Value
                If Range("D" & check_row).Value <> "" Then
                        Max_Bit = Range("B" & check_row).Value
                        check_row = check_row + 1
                End If
        Loop
End Function

Private Function Check_Values_Are_Clean() As Boolean
        'Makes sure some idiot didn't use text values instead of numbers and breaks everything.
        Check_Values_Are_Clean = True
        Dim check_row As Long
        check_row = 7
        
        If IsNumeric(Range("Target_Value").Value) Then
        Else
                Check_Values_Are_Clean = False
                Exit Function
        End If
        
        Do While Check_Values_Are_Clean And Range("D" & check_row).Value <> ""
                DoEvents
                If IsNumeric(Range("D" & check_row).Value) Then
                Else
                        Check_Values_Are_Clean = False
                End If
                check_row = check_row + 1
        Loop
        
        
End Function














