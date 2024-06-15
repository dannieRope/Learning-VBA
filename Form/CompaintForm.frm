VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CompaintForm 
   Caption         =   "Customer Complaint Form"
   ClientHeight    =   13185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10245
   OleObjectBlob   =   "CompaintForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CompaintForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Click()
CompaintForm.Show
End Sub

Private Sub UserForm_Initialize()
'Keep NameBox empty by default
NameTextBox.Value = ""

'Keep AgecomboBox clear by Default
AgeComboBox.clear
'Add age values to the age combo
With AgeComboBox
       .AddItem "18-25"
       .AddItem "25-35"
       .AddItem "35-45"
       .AddItem "45-55"
       .AddItem "55-65"
       .AddItem "65>"
       
End With

'Keep the Male the male and female optionbutton clear by default
MaleOptionButton.Value = False
FemaleOptionButton.Value = False

'Keep the email text box empty by default
EmailTextBox = ""

'Keep both check boxes unchecked by default
YesCheckBox.Value = False
NoCheckBox.Value = False

'Add items to listbox for countries
With CountryListBox
         .AddItem "USA"
         .AddItem "UK"
         .AddItem "India"
         .AddItem "Nigeria"
         .AddItem "Ghana"
         .AddItem "Spain"
         .AddItem "China"
         .AddItem "Benin"
         .AddItem "Togo"
         .AddItem "Germany"
End With

'keep the billtextbox button empty

BillTextBox.Value = ""

'set focus on NameTextBox
'NameTextBox.SetFocus

End Sub

Private Sub ClearCommandButton_Click()
'Set the form to default

Call UserForm_Initialize

End Sub

Private Sub CancelCommandButton_Click()
'Hide or close the form

Unload Me

End Sub
Private Sub BillSpinButton_Change()

BillTextBox.Text = BillSpinButton.Value

End Sub

Private Sub OkayCommandButton_Click()

Application.ScreenUpdating = False

Dim newemptyrow As Long
'Make the sheet active
Sheets("Userform").Activate

'Determine the empty row
newemptyrow = WorksheetFunction.CountA(Range("A:A")) + 1

Cells(newemptyrow, 1).Value = NameTextBox.Value

Cells(newemptyrow, 2).Value = AgeComboBox.Value

If FemaleOptionButton.Value = True Then
     Cells(newemptyrow, 3) = "Female"
Else
    Cells(newemptyrow, 3) = "Male"
End If

Cells(newemptyrow, 4).Value = BillTextBox.Value

Cells(newemptyrow, 5).Value = CountryListBox.Value

Cells(newemptyrow, 6).Value = EmailTextBox.Value

If YesCheckBox.Value = True Then
    Cells(newemptyrow, 7).Value = "Yes"
Else: Cells(newemptyrow, 7).Value = "No"
End If

Unload Me


End Sub


