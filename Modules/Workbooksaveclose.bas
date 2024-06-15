Attribute VB_Name = "Workbooksaveclose"
Sub Savecloseworkbook()
Workbooks("VBA_extra").Sheets(1).Range("A1:A6") = "VBA Extra"
ThisWorkbook.Save
Workbooks("VBA_extra").Close

End Sub
