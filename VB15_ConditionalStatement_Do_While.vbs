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

' Do . . . While. . .

' Loops - Do While, Do Until
' 		- wscript.quit - completely quit the script
' 		- exit do - exiting in the do loop then run the  remaining codes outside the do loop

' ============================================================

option explicit
' on error resume next

Dim a 
Dim b
Dim c
a = 0
b = 0
c = 0

Do
	a = a + 1
	MsgBox a & ": I am inside the loop"

	If a = 5 Then
		exit do
		' wscript.quit
	End If
Loop

MsgBox "I am Outside the loop"


'=======================================
' Exit when TRUE

Do Until b = 6
	MsgBox "b = " & b
	b = b + 1
Loop

MsgBox "I am Outside the loop Again"



'=======================================
' Exit when FALSE

Do While c < 6
	MsgBox "c = " & c
	c = c + 1
Loop

MsgBox "I am Outside the loop AGAIN"