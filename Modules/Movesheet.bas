Attribute VB_Name = "Movesheet"
Sub movesheet()
Sheets("VBA").Move After:=Sheets("clear")
Sheets("VBA").Move Before:=Sheets("Clear")
End Sub
