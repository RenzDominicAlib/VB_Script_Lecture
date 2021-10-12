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

' If . . . condition . . .

' ============================================================

option explicit
' on error resume next

'declaration of variables and constant

Const BOX_TITLE = "Student Score" 

Dim a
Dim b

a = 8
b = 6


If a > b Then	
MsgBox "a is greater than b"

ElseIf a = b Then
MsgBox "a is equal to b"

Else 
MsgBox "a is less than b"
End If
