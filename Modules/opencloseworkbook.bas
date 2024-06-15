Attribute VB_Name = "opencloseworkbook"
Sub opencloseworkbook()
Workbooks.Open Filename:="C:\Users\HP\Desktop\excel\VBA_extra.xlsx"
Workbooks("VBA_extra").Sheets(1).Range("B1:B6") = "Excel"
Workbooks("VBA_extra").Save
Workbooks("VBA_extra").Close

End Sub
