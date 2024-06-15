Attribute VB_Name = "columnWRowH"
Sub column_width()
Range("A1:B10").ColumnWidth = 15
Range("A1:B10").Columns.ColumnWidth = 25
Range("A1:B10").Columns.AutoFit
Range("A1:B10").RowHeight = 6
Range("A1:B10").Rows.RowHeight = 4
Range("A1:B10").Rows.AutoFit
End Sub
