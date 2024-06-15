VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CountryState 
   Caption         =   "Country and State"
   ClientHeight    =   4230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8715.001
   OleObjectBlob   =   "CountryState.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CountryState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
'add countries to the combo list
With CountryComboBox
     .AddItem "Ghana"
     .AddItem "Nigeria"
     .AddItem "Togo"
End With
'clear statecomboBox at default
StateComboBox.clear
End Sub

Private Sub CountryComboBox_Change()
Dim CountryName As String

CountryName = CountryComboBox.Value

Select Case CountryName
     Case Is = "Ghana"
         With StateComboBox
            .AddItem "Greater Accra"
            .AddItem "Volta Region"
            .AddItem "Western Region"
            .AddItem "Asante Region"
         End With
         
         
     Case Is = "Nigeria"
         With StateComboBox
             .AddItem "Oyo State"
             .AddItem "Edo State"
             .AddItem "Anambra State"
             .AddItem "Lagos State"
         End With
     Case "Togo":
           With StateComboBox
              .AddItem "Kpalime"
              .AddItem "Asigame"
              .AddItem "Lome"
           End With
End Select

End Sub
Private Sub OkayCommandButton_Click()
Application.ScreenUpdating = False

Sheets("Country").Activate
Dim emptyrow As Integer

emptyrow = WorksheetFunction.CountA(Range("A:A")) + 1
Cells(emptyrow, 1) = CountryComboBox.Value
Cells(emptyrow, 2) = StateComboBox.Value

Unload Me

End Sub


Private Sub CancelCommandButton_Click()

Unload Me


End Sub




