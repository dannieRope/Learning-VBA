Attribute VB_Name = "messagebox"
Sub msgbox1()
MsgBox "Hello World"
MsgBox "Welcome to my first VBA tutorial"
End Sub

Sub msgbox2()
MsgBox "Hello World", 1 'okay, cancel
MsgBox "Hello World", 2 'Try,abort, ignore
MsgBox "Hello World", 3 'yes,no,cancel
MsgBox "Hello World", 4 'yes,no
MsgBox "Hello World", vbOKCancel
End Sub

Sub msgbox3()
MsgBox "Hello World", 16 'stop
MsgBox "Hello World", 32 'Help
MsgBox "Hello World", 48 'Warning
MsgBox "Hello World", 64 'Information
MsgBox "Hello World", vbOKCancel 'cancel

MsgBox "Hello World", vbCritical 'stop
MsgBox "Hello World", vbQuestion 'Help
MsgBox "Hello World", vbExclamation 'Warning
MsgBox "Hello World", vbInformation 'Information
MsgBox "Hello World", vbOKCancel 'cancel
End Sub

Sub msgbox4()
Dim a As Integer
a = MsgBox("Hello World", 1)
MsgBox a
If a = 1 Then
    MsgBox "Thanks for pressing ok"
Else: MsgBox "You pressed cancel"
End If

End Sub

Sub msgbox5()
MsgBox "Hello! Welcome", 1, "VBA Tutorial"
End Sub



