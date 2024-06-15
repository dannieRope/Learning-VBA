Attribute VB_Name = "CopySheet"
Sub copysheet()
Sheets("VBA").Copy After:=Sheets("Clear")
Sheets("Clear").Copy Before:=Sheets("Right")
End Sub
