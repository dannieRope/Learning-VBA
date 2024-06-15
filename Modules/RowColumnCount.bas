Attribute VB_Name = "RowColumnCount"
Sub rowcolumncount()
Dim x As Long
Dim y As Long
Dim i As Integer
Dim z As Integer
Dim rng As Range

Set rng = Range("A1:D11")
x = Rows.Count
y = Columns.Count
i = rng.Rows.Count
z = rng.Columns.Count
Debug.Print x
Debug.Print y
Debug.Print i
Debug.Print z

End Sub
