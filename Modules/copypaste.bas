Attribute VB_Name = "copypaste"
Sub copypaste()
'First method of coping and pasting values
[C2:B10] = [B2:B10].Value
[B2:B10].Copy
[D2:D10].pasteSpecial
Application.CutCopyMode = False
End Sub
