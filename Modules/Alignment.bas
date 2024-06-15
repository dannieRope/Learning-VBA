Attribute VB_Name = "Alignment"
Sub alignment()
'name cell one as gender
Range("A1").Value = "Gender"
Cells(1, 1).HorizontalAlignment = xlLeft
Cells(1, 1).HorizontalAlignment = xlRight
Cells(1, 1).HorizontalAlignment = xlCenter

Cells(1, 1).VerticalAlignment = xlTop
Cells(1, 1).VerticalAlignment = xlBottom
Cells(1, 1).VerticalAlignment = xlCenter
End Sub
