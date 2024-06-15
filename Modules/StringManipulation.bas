Attribute VB_Name = "StringManipulation"
Sub stringManipulation()
Dim rng As Range
Dim cell As Range

Set rng = Range("A2:A8")
For Each cell In rng
   cell.Offset(0, 1).Value = Left(cell, 3)
   cell.Offset(0, 2).Value = Right(cell, 3)
Next
End Sub

Sub stringManipulation2()
Dim x As Integer
For x = 2 To 8

End Sub
