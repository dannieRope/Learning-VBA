Attribute VB_Name = "StaticArray"
Sub IdArrary()
Dim Marks(4) As Integer
Marks(1) = 10
Marks(2) = 20
Marks(3) = 30
Marks(4) = 40

For i = 0 To 4
  Debug.Print Marks(i)
Next
End Sub

Sub twodArray()
Dim a(2, 2) As Integer
a(0, 0) = 10
a(0, 1) = 20
a(0, 2) = 30

a(1, 0) = 40
a(1, 1) = 50
a(1, 2) = 60

a(2, 0) = 70
a(2, 1) = 80
a(2, 2) = 90

For i = 0 To 2
    For e = 0 To 2
      Debug.Print a(i, e)
    Next
Next

End Sub

Sub splitecg()
Dim arr() As String
Dim x As String

x = "First,Second,Third"
arr = Split(x, ",")

Debug.Print arr(0), arr(1), arr(2), arr(3)

End Sub

Sub redimsepcial()
Dim number() As Integer
ReDim number(3)
number(0) = 10
number(1) = 20
number(2) = 30
number(3) = 40

ReDim Preserve number(5)
number(4) = 50
number(5) = 60

For i = LBound(number) To UBound(number)
    Debug.Print number(i)
Next

End Sub

Sub redimspecial()
Dim r(2, 2)
r(0, 0) = 1
r(0, 1) = 2
r(0, 2) = 3
r(1, 0) = 4
r(1, 1) = 5
r(1, 2) = 6
r(2, 0) = 7
r(2, 1) = 8
r(2, 2) = 9

For i = 0 To 2
    For e = 0 To 2
      Debug.Print r(i, e)
    Next
Next


End Sub
