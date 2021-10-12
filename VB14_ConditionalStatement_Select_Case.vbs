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

' Select. . .Case. . .Condition. . .

' <> unequal

' ============================================================

option explicit
' on error resume next

'declaration of variables and constant

Const BOX_TITLE = "Student Score" 
Dim strFname
Dim strLname
Dim intId
Dim strBirthmonth
Dim strLCaseBirthmonth
Dim strMessage

'process or actual flow

strFname = InputBox("Student First Name: ")
strLname = InputBox("Student Last Name: ")
intId = InputBox("Student ID number: ", BOX_TITLE, "Enter 5-Digits ID here")
strBirthmonth = InputBox("Birth Month: ")
strLCaseBirthmonth = LCase(strBirthmonth)

Select Case strLCaseBirthmonth
	Case "january"
	strMessage = "January Babies are Handsome!"
	Case "february"
	strMessage = "february Babies are Handsome!"
	Case "march"
	strMessage = "march Babies are Handsome!"
	Case "april"
	strMessage = "april Babies are Handsome!"
	Case "may"
	strMessage = "may Babies are Handsome!"
	Case "june"
	strMessage = "June Babies are Handsome!"
	Case "july"
	strMessage = "July Babies are Handsome!"
	Case "august"
	strMessage = "august Babies are Handsome!"
	Case "september"
	strMessage = "September Babies are Handsome!"
	Case "october"
	strMessage = "October Babies are Handsome!"
	Case "november"
	strMessage = "November Babies are Handsome!"
	Case "december"
	strMessage = "December Babies are Handsome!"
	Case Else
	strMessage = "Invalid Birth Month, Try again!"
End Select


MsgBox "Student Full Name : " & strFname & " " & strLname & vbLf &_
"ID number: " & intId, 0, BOX_TITLE
MsgBox strMessage


