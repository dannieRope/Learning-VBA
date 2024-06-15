Attribute VB_Name = "WithBlockfont"
Sub withblock()
[A4:A10] = "Women"

With [A4:A10].Font
     .Name = "Italic"
     .Size = 12
     .Bold = False
     .Underline = True
End With

End Sub
