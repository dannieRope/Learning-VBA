Attribute VB_Name = "dowhileloop"
Sub dowhileloop()
Dim i As Integer
i = 1
Do While Cells(i, 1).Value <> ""
 Cells(i, 2).Value = Cells(i, 1) * 2
 i = i + 1
Loop
   

End Sub
