Attribute VB_Name = "Addsheet"
Sub addsheet()
Sheets.Add
Worksheets.Add
Sheets.Add Before:=Sheets("Clear")
Worksheets.Add After:=Worsheets("VBA")
End Sub
