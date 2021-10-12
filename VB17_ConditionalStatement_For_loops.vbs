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
'UBound - Highest index of array; upper bound is index 4 which is equal to 5



' ============================================================

option explicit
' on error resume next

Const BOX_TITLE = "For Loop" 

Dim total, arrNums, i

total = 0
arrNums = Array(1, 2, 3, 4, 5)

For i = UBound(arrNums) To LBound(arrNums) Step -1
	total = total + arrNums(i)

	'first: 0 = 0 + arrNums(4) ---> 5 = 0 + 5
	'second: 5 = 5 + arrNums(3) ---> 9 = 5 + 4
	'third: 9 = 9 + arrNums(2) ---> 12 = 9 + 3
	'fourth: 12 = 12 + arrNums(1) ---> 14 = 12 + 2
	'fifth: 14 = 14 + arrNums(0) ---> 15 = 14 + 1
	
	If total > 10 Then
		MsgBox "If - Total value: " & total & "; array value: " & arrNums(i), 0, BOX_TITLE
	Else
		MsgBox "Else - Total value: " & total & "; array value: " & arrNums(i), 0, BOX_TITLE
	End If
Next

MsgBox "Outside the for loop - Total value: " & total, 0, BOX_TITLE