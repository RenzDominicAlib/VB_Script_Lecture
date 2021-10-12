' ============================================================
'Using Built in function with strings, MID, LEN


' ============================================================

' Using ByRef and ByVal with parameters/arguments

' ByVal ==> Passed by value
' ByRef ==> Passed by value or reference
' =============================================================

' option explicit
' on error resume next

Dim MyDate
MyDate = Date 

Dim month, day, year
day = mid(MyDate, 1, 2)
month = mid(MyDate, 4, 2)
year = mid(MyDate, 7, 4)

Sub DisplayMsg (strMsg, intResult)
	MsgBox strMsg & " : " & intResult, 0, BOX_TITLE
End Sub

Call DisplayMsg("Example 1", MyDate)
Call DisplayMsg("Example 2", Date)
Call DisplayMsg("Example 3", month)
Call DisplayMsg("Example 4", day)
Call DisplayMsg("Example 5", year)

'Using AddDate

Dim MyNewDate
MyNewDate = DateAdd("m", 1, MyDate)  'Adding 1 month
Call DisplayMsg("Example 6", MyNewDate)


MyNewDate2 = DateAdd("m", 8, "15-January-1995")  'Adding 8 months
Call DisplayMsg("Example 7", MyNewDate2)


MyNewDate3 = DateAdd("d", 2, "15-January-1995")  'Adding 2 days
Call DisplayMsg("Example 8", MyNewDate3)

'Extracting hour, minute, and Second from time

Call DisplayMsg("Example 9", Time)

Dim MyTime, MyHour, MyMinute, MySecond
MyTime = Time 
MyHour = Hour(MyTime)
MyMinute = Minute(MyTime)
MySecond = Second(MyTime)

Call DisplayMsg("Example 10", MyHour)
Call DisplayMsg("Example 11", MyMinute)
Call DisplayMsg("Example 12", MySecond)

'Using Now (get current date and time)
Dim nowTime
nowTime = now 
Call DisplayMsg("Example 13", nowTime)
