' ============================================================
				
' 				Header, comments, dates, important notes

' 1. ( vbLf ) - is used to declare a newline.
' 2. ( &_ ) - is use to break line in a code editor .
' 3. ( option explicit ) - used to strict the code in declaring variables.
' 4. ( on error resume next ) - used to continue the running/execution disregarding the error.
' 5. ( InputBox() ) - used to create a inputbox.
' 6. ( CInt() ) - converting into the real integers


' ============================================================

' Sub Procedure --- Does not return a value
' Function Procedure --- can return a value



' ============================================================

option explicit
' on error resume next

Const BOX_TITLE = "Box Title" 
Dim input1
Dim input2
Dim total
Dim sum



Function Add(num1, num2)
	sum = num1 + num2
	Add = sum
End Function

input1 = InputBox("Input one: ", BOX_TITLE, "Enter valid number here")
input2 = InputBox("Input two: ", BOX_TITLE, "Enter valid number here")
total = Add(CInt(input1), CInt(input2))
MsgBox "The Sum of two numbers is " & total