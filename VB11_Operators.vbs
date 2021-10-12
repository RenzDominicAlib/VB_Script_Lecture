' ============================================================
				
' 				Header, comments, dates, important notes

' 1. ( vbLf ) - is used to declare a newline.
' 2. ( &_ ) - is use to break line in a code editor .
' 3. ( option explicit ) - used to strict the code in declaring variables.
' 4. ( on error resume next ) - used to continue the running/execution disregarding the error.
' 5. ( InputBox() ) - used to create a inputbox.
' 6. ( cint() ) - converting into the real integers
' 7. ( cdbl() ) - converting into the real double
' 8. ( cdate() ) - converting into the real date
' 9. ( cbool() ) - converting into the real boolean
' 10. ( cstr() ) - converting into the real string
' 10. ( ccur() ) - converting into the real currency
' 10. ( cbyte() ) - converting into the real Byte
' 10. ( Isnumeric() ) - boolean validation
' 10. ( Isempty() ) - boolean validation
' 10. ( Isarray() ) - boolean validation
' 10. ( Isobject() ) - boolean validation



' ============================================================

'Arithmetic Operators : ^, -, *, /, \, Mod, +, -, &
'Comparison           : =, <>, <, >, <=, >=, Is
'Logical              : Not, And, Or, Xor, Eqv, Imp

' ============================================================

option explicit
' on error resume next

'declaration of variables and constant

Const BOX_TITLE = "Student Score" 
Dim strFname
Dim strLname
Dim intId
Dim intInputOne
Dim intInputTwo
Dim intProduct
Dim intPower
Dim intModulus
Dim intAverage

'process or actual flow

strFname = InputBox("Student First Name: ")
strLname = InputBox("Student Last Name: ")
intId = InputBox("Student ID number: ", BOX_TITLE, "Enter 5-Digits ID here")
intInputOne = InputBox("First Input: ", BOX_TITLE, "Enter Valid number", 5000,5000)
intInputTwo = InputBox("Second Input: ", BOX_TITLE, "Enter Valid number", 5000, 5000)

If isnumeric(intInputOne) and isnumeric(intInputTwo) then

intAverage = (CInt(intInputOne) + CInt(intInputTwo)) / 2
intProduct = CInt(intInputOne) * CInt(intInputTwo)
intPower = CInt(intInputOne) ^ CInt(intInputTwo)
intModulus = CInt(intInputOne) Mod CInt(intInputTwo)
Else 
MsgBox "Enter integer value!"
End if

MsgBox "Student Full Name : " & strFname & " " & strLname & vbLf &_
"ID number: " & intId, 0, BOX_TITLE
MsgBox "Average Score is : " & intAverage
MsgBox "Product is : " & intProduct
MsgBox "Power Score is : " & intPower
MsgBox "Modulus is : " & intModulus