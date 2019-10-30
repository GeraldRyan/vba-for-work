Attribute VB_Name = "ChangeToPercent"
Sub changetopercent()
'
' changetopercent Macro

Dim Billy As range, john As range
Set Billy = Selection

For Each john In Billy
If Not IsEmpty(john) Then
    john = john / 100
    john.Style = "Percent"  'Style is a built-in function
End If
Next
End Sub


Sub ParenthesesWithNegativePercent()


Application.Selection.NumberFormat = "0.0%;(0.0%)"
End Sub






















Sub change2percent2()
Dim n1 As range
Dim n2 As range

    Set n1 = Application.InputBox(Prompt:= _
                    "Select cells to create formula", _
                    Title:=sTitle & " Creator", Type:=8)
        Set n2 = Application.InputBox(Prompt:= _
                    "Select cells to create formula", _
                    Title:=sTitle & " Creator", Type:=8)
                    
                    
                    
End Sub
