Attribute VB_Name = "PasteSpecial"
Sub pasteSpecial()
[e2:e10] = "Boy"
[e2:e10].Copy
[f2:f10].pasteSpecial xlPasteFormats
[f2:f10].pasteSpecial xlPasteColumnWidths
[f2:f10].pasteSpecial xlPasteValues
Application.CutCopyMode = False
End Sub
