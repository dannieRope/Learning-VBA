Attribute VB_Name = "Delete"
Sub delete()
[A5].delete 'delete a cell
[A1:B4].delete 'delete a range
[A5].EntireRow.delete
[A:A].EntireColumn.delete
End Sub
