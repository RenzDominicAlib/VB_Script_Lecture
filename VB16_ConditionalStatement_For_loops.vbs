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

' For . . . While. . .
' By default it incremented by 1
' Step 2 means incremented by 2
'UBound - Highest index of array;



' ============================================================

option explicit
' on error resume next

Dim a 
Dim b

Dim x
Dim y
Dim myloop

For a = 1 To 10
MsgBox a & " : I am inside the for loop"
Next

MsgBox "OUTSIDE"

'=======================================
' Step or incremented by 2

For b = 1 To 10 Step 2
MsgBox b & " : I am inside the for loop with increment of 2"
Next

MsgBox "OUTSIDE"


'=======================================
' Decrementing using negative number 1

x = 10
y = 5

For myloop = x To y Step -1
MsgBox myloop & " : I am inside the for loop with decrement of 1"
Next

MsgBox "OUTSIDE"


'=======================================
' Decrementing using negative number 2

x = 10
y = 5

For myloop = x To y Step -2
MsgBox myloop & " : I am inside the for loop with decrement of 2"
Next

MsgBox "OUTSIDE"