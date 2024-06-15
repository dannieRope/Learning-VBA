Attribute VB_Name = "Withblockborder"
Sub border()
Range("A12:C16") = "No"

With Range("A12:C16").Borders
     .Color = vbGreen
     .Weight = 3
     .LineStyle = xlDash
End With

End Sub
