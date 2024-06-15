Attribute VB_Name = "loop1"
Sub loop1()
Dim x As Integer
  For x = 1 To 10
     MsgBox x 'Display as message
Next
End Sub

Sub loop2()
Dim x As Integer
 For x = 1 To 10 Step 2
   Cells(x, 1).Value = x 'Enter data in cell
   Next
End Sub

Sub loop3()
Dim x As Integer
For x = 1 To 56
 Cells(x, 1).Value = x 'Enter number in to cells
 Cells(x, 2).Interior.ColorIndex = x 'Color cells
 Next
End Sub

Sub loop4()
Dim x As Integer
For x = 20 To 1 Step -1
  Cells(x, 3) = x
  Next
End Sub

Sub loop5()
Dim x As Integer
For x = 1 To 10
    Cells(x, x).Value = x
Next x
End Sub

Sub loop6()
Dim x As Integer
For x = 1 To ThisWorkbook.Sheets.Count
  MsgBox ThisWorkbook.Sheets(x).Name
Next
End Sub

Sub loop7()
Dim sht As Worksheet
For Each sht In ThisWorkbook.Sheets
     MsgBox sht.Name
Next
End Sub


