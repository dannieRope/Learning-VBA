Attribute VB_Name = "DoUntilLoop3"
Sub dountilloop()
Dim x As Integer
x = 1
Do Until x = 20
Cells(x, 3).Value = x
x = x + 1
Loop

End Sub
