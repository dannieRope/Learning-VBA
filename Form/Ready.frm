VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ready 
   Caption         =   "Ready????"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8685.001
   OleObjectBlob   =   "Ready.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Ready"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub okCommandButton_Click()
QFrame.Visible = False
End Sub

Private Sub UserForm_Activate()
QFrame.Visible = False
End Sub


Private Sub ReadyCommandButton_Click()

QFrame.Visible = True

End Sub


