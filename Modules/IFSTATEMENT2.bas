Attribute VB_Name = "IFSTATEMENT2"
Sub ifstatement()
If Range("E1").Value = 200 Then Range("F1").Value = "Large"

If Range("E4").Value >= 100 Then Range("F4").Value = "Pass"

End Sub

Sub iistatement2()
If Cells(1, 6).Value = "Large" Then Cells(1, 7) = "London" Else Cells(1, 7).Value = "USA"
If Cells(2, 6).Value = "Pass" Then Cells(2, 7) = "Correct" Else Cells(2, 7).Value = "Wrong"

End Sub

Sub ifstatement3()
If Cells(3, 6).Value = "Small" Then
    Cells(3, 7) = "India"
ElseIf Cells(3, 6).Value = "Large" Then
    Cells(3, 7) = "London"
ElseIf Cells(3, 6).Value = "Pass" Then
    Cells(3, 7) = "USA"
Else:
    Cells(3, 7).Value = "Nothing"
End If

End Sub

Sub ifstatementwithloop()

Dim rng As Range
Dim cell As Range

Set rng = Range("B2:B11")

For Each cell In rng
    If cell.Value <= 40 Then
       cell.Offset(0, 1).Value = "Fail"
    Else: cell.Offset(0, 1).Value = "Pass"
End If
Next

End Sub

Sub ifstatementloop2()
Dim x As Integer
For x = 2 To 11
  If Cells(x, 2).Value <= 40 Then
        Cells(x, 4).Value = "Fail"
        Cells(x, 4).Interior.Color = vbRed
  Else: Cells(x, 4).Value = "Pass"
        Cells(x, 4).Interior.Color = vbGreen
  End If
Next
End Sub
