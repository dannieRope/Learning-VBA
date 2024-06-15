VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Subject 
   Caption         =   "Favorite Subject"
   ClientHeight    =   4020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7260
   OleObjectBlob   =   "Subject.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Subject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelCommandButton_Click()
Unload Me

End Sub

Private Sub okCommandButton_Click()
If EnglishCheckBox.Value = True Then
    ThisWorkbook.Sheets("Subject").Range("English_Range") = ThisWorkbook.Sheets("Subject").Range("English_Range") + 1
End If


If MathsCheckBox.Value = True Then
    ThisWorkbook.Sheets("Subject").Range("Maths_Range") = ThisWorkbook.Sheets("Subject").Range("Maths_Range") + 1
End If


If ScienceCheckBox.Value = True Then
    ThisWorkbook.Sheets("Subject").Range("Science_Range") = ThisWorkbook.Sheets("Subject").Range("Science_Range") + 1
End If

EnglishCheckBox.Value = False
MathsCheckBox.Value = False
ScienceCheckBox.Value = False

End Sub
