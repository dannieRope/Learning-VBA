Attribute VB_Name = "SheetRef"
Sub sheetref()
Range("A1:A5") = "Excel VBA"
Sheets(4).Range("A1:A5") = "Excel VBA"
Sheets("love").Range("A6:A10") = "VBA love"
Worksheets("Clear").Range("A6:A10") = "VBA love"
End Sub
