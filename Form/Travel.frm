VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Travel 
   Caption         =   "Travel Quesion"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7860
   OleObjectBlob   =   "Travel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Travel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Click()
'Untick all buttons as default

AOptionButton.Value = False
BOptionButton.Value = False
COptionButton.Value = False

End Sub
Private Sub SubmitCommandButton_Click()

'Activate the appropriate sheet
Sheets("Travel").Activate

Dim emptyrow As Integer
'Identity the next available empty cell

emptyrow = WorksheetFunction.CountA(Range("A:A")) + 1


If AOptionButton.Value = True Then
    Cells(emptyrow, 1) = "Bus"
    MsgBox "you have selected" & " " & Cells(emptyrow, 1).Value
ElseIf BOptionButton.Value = True Then
    Cells(emptyrow, 1) = "Car"
    MsgBox "you have selected" & " " & Cells(emptyrow, 1).Value
ElseIf COptionButton.Value = True Then
    Cells(emptyrow, 1) = "Flight"
    MsgBox "you have selected" & " " & Cells(emptyrow, 1).Value
Else: MsgBox "Choose An Answer"

End If



Unload Me


End Sub
