Attribute VB_Name = "SelectCase"
Sub selectcase()
Dim Var As Integer
Var = inputbox("Enter Month Number")

Select Case Var
    Case 1: MsgBox "Month is January"
    Case 2: MsgBox "Month is Febuary"
    Case 3: MsgBox "Month is March"
    Case 4: MsgBox "Month is April"
    Case 5: MsgBox "Month is May"
    Case 6: MsgBox "Month is June"
    Case 7: MsgBox "Month is July"
    Case 8: MsgBox "Month is August"
    Case 9: MsgBox "Month is September"
    Case 10: MsgBox "Month is October"
    Case 11: MsgBox "Month is Novermber"
    Case 12: MsgBox "Month is December"
    Case Else: MsgBox "Invalid Month Number"
End Select

End Sub

Sub selectcase2()
Select Case Cells(2, 4).Value
     Case "Fail": Cells(2, 5).Value = "Correct"
     Case Else: Cells(2, 5).Value = "Wrong"
End Select


End Sub
