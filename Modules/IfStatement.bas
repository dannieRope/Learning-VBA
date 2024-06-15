Attribute VB_Name = "IfStatement"
Sub ifstatement()

    Dim cell As Range
    Dim rng As Range
    
    ' Define the range to be checked
    Set rng = Range("E1:E9")
    
    ' Loop through each cell in the range
    For Each cell In rng
        If cell.Value <= 20 Then
            ' Set the corresponding cell in column F to "Small"
            cell.Offset(0, 1).Value = "Small"
        End If
    Next cell
End Sub






