Attribute VB_Name = "HandlingError"
Sub handleError()
On Error Resume Next
MsgBox "Accra We Dey"
MsgBox 10
MsgBox 10 / 0
MsgBox "Here we go"

End Sub

Sub handleError1()
On Error GoTo abc
    MsgBox 10 / 0

Done:
    Exit Sub
abc:
    MsgBox "You can't divide by zero"
   
End Sub
