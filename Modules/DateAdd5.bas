Attribute VB_Name = "DateAdd5"
Sub dateadd1()
Dim mydate As Date
Dim y As Date
Dim x As Date
Dim z As Date


mydate = #2/14/2024#

y = DateAdd("yyyy", 1, mydate) 'add one year
x = DateAdd("m", 1, mydate) 'add one month
z = DateAdd("d", 1, mydate) 'add one day

i = DateAdd("h", 2, "30 Dec 2024 11:45:09") 'add 2 hours
e = DateAdd("n", 30, "30 Dec 2024 11:45:09") 'add 30 mins
q = DateAdd("s", 1, "30 Dec 2024 11:45:09") 'add 1 sec

Debug.Print mydate, y, x, z, i, e, q

End Sub

Sub dateadd5()
    Dim mydate As Date
    Dim y As Date
    Dim x As Date
    Dim z As Date
    
    mydate = #2/14/2024#
    y = DateAdd("yyyy", 1, mydate) 'add one year
    x = DateAdd("m", 1, mydate) 'add one month
    z = DateAdd("d", 1, mydate) 'add one day

    Debug.Print y
    Debug.Print x
    Debug.Print z
End Sub

Sub datepart1()
mydate = #1/2/2024#
x = DatePart("yyyy", mydate)
y = DatePart("m", mydate)
z = DatePart("d", mydate)
i = DatePart("q", mydate)
u = Format(mydate, "mmm")

Debug.Print x, y, z, i, u

End Sub

Sub datepart2()
mydate = #1/2/2024#
x = Month(mydate)
y = Day(mydate)
z = Year(mydate)
Debug.Print x, y, z, i

End Sub

Sub datepart3()
Dim x As Variant
Dim y As Variant

x = "2014-02-01"
y = "30 Dec 2024"
i = CDate(x)
z = CDate(y)

Debug.Print x
Debug.Print y
Debug.Print z
Debug.Print i

End Sub

Sub dateparttime()
y = Date
x = Now()
z = Time()

h = Hour(z)
m = Minute(z)
s = Second(z)

Debug.Print x, y, z
Debug.Print h, m, s


End Sub

Sub timeserial1()
Debug.Print TimeSerial(4, 50, 20)
Debug.Print TimeSerial(21, 45, 1)
Debug.Print TimeValue("1:45:20")
Debug.Print TimeValue("21:45:30")


End Sub
