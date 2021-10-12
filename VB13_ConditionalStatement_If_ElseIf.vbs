' ============================================================
				
' 				Header, comments, dates, important notes

' 1. ( vbLf ) - is used to declare a newline.
' 2. ( &_ ) - is use to break line in a code editor .
' 3. ( option explicit ) - used to strict the code in declaring variables.
' 4. ( on error resume next ) - used to continue the running/execution disregarding the error.
' 5. ( InputBox() ) - used to create a inputbox.
' 6. ( CInt() ) - converting into the real integers


' ============================================================

'Arithmetic Operators : ^, -, *, /, \, Mod, +, -, &
'Comparison           : =, <>, <, >, <=, >=, Is
'Logical              : Not, And, Or, Xor, Eqv, Imp

' If . . .ElseIf. . .Else. . . condition . . .

' <> unequal

' ============================================================

option explicit
' on error resume next

'declaration of variables and constant

Const BOX_TITLE = "Student Score" 
Dim strFname
Dim strLname
Dim strUname
Dim strCUname
Dim intId
Dim intInputOne
Dim intInputTwo
Dim intInputThree
Dim intAverage
Dim strComment

'process or actual flow

strFname = InputBox("Student First Name: ")
strLname = InputBox("Student Last Name: ")
strUname = InputBox("User name: ")
strCUname = InputBox("Confirm User name: ")
intId = InputBox("Student ID number: ", BOX_TITLE, "Enter 5-Digits ID here")
intInputOne = InputBox("First Attempt Score: ", BOX_TITLE, "Enter Valid number/score", 5000,5000)
intInputTwo = InputBox("Second Attempt Score: ", BOX_TITLE, "Enter Valid number/score", 5000, 5000)
intInputThree = InputBox("Third Attempt Score: ", BOX_TITLE, "Enter Valid number/score", 5000, 5000)
intAverage = (CInt(intInputOne) + CInt(intInputTwo) + CInt(intInputThree)) / 3


If intAverage >= 90 And intAverage <= 100 Then
	strComment = "Excellent! You've passed the Test!"

ElseIf intAverage <= 89 And intAverage >= 75 Then
	strComment = "Very Good! You've passed the Test!"

ElseIf intAverage <= 74 And intAverage >= 0 Then
	strComment = "Sorry! You've failed the Test!"

Else 
	strComment = "You've input invalid scores"
End If


If strUname <> strCUname Then
	MsgBox "User Name Confirmation did not Match"
Else
	MsgBox "Student Full Name : " & strFname & " " & strLname & vbLf &_
	"ID number: " & intId, 0, BOX_TITLE
	MsgBox "Average Score is : " & intAverage
	MsgBox strComment
End If

