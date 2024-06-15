Attribute VB_Name = "getworkbookname"
Sub getworkbookname()
MsgBox (ActiveWorkbook.Name)
MsgBox (ThisWorkbook.Name)
Workbooks("VBA_Extra").Activate
End Sub
