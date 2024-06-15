Attribute VB_Name = "Backgroundcolor"
Sub BackgroundColor()
    ' Set the background color for all cells in the active sheet to RGB color (149, 163, 164)
    Cells.Interior.Color = RGB(149, 163, 164)
    Range("A12:C16").Interior.Color = vbWhite
    Cells.Interior.Color = vbWhite
End Sub

