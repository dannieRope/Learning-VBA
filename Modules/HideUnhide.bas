Attribute VB_Name = "HideUnhide"
Sub hideunhide()
Range("A:A").Columns.Hidden = True
Range("A:A").Columns.Hidden = False
Range("A1").Rows.Hidden = True
Range("A1").Rows.Hidden = False


End Sub
