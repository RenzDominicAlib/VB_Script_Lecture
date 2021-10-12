' ============================================================
				
' 				Header, comments, dates, important notes

' 1. ( vbLf ) - is used to declare a newline.
' 2. ( &_ ) - is use to break line in a code editor .
' 3. ( option explicit ) - used to strict the code in declaring variables.
' 4. ( on error resume next ) - used to continue the running/execution disregarding the error.
' 5. ( InputBox() ) - used to create a inputbox.
' 6. ( CInt() ) - converting into the real integers


' ============================================================

' Using ByRef and ByVal with parameters/arguments

' ByVal ==> Passed by value
' ByRef ==> Passed by value or reference
' =============================================================

Dim mynum1, mynum2, result
mynum1 = 5
mynum2 = 7

result = AddNumbers((mynum1), mynum2)  '<========= additional parenthesis means byVal

MsgBox "M1 : Result is " & result
MsgBox "M2 : mynum1 = " & mynum1 & " and mynum2 = " & mynum2


Function AddNumbers(num1, num2)			'<======ByRef, Default
	AddNumbers = num1 + num2
	num1 = 11111
	num2 = 22222
End Function


