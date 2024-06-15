Attribute VB_Name = "Dowhile2"
Sub dowhile2()
Dim i As Integer
i = 1
Do While i <= 10
Cells(i, 1).Value = i
Cells(i, 2).Interior.ColorIndex = i
i = i + 1
Loop

End Sub
