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
' Calculator code


' ============================================================

option explicit
' on error resume next

Const BOX_TITLE = "Box Title" 
Dim input1, input2, result, intInputOne, intInputTwo


Function Addition(num1, num2)
	Addition = num1 + num2
End Function

Function Subtraction(num1, num2)
	Subtraction = num1 - num2
End Function

Function Multiplication(num1, num2)
	Multiplication = num1 * num2
End Function

Function Division(num1, num2)
	Division = num1 / num2
End Function

Sub DisplayMsg (strMsg, intResult)
	MsgBox strMsg & " : " & intResult, 0, BOX_TITLE
End Sub



input1 = InputBox("Input one: ", BOX_TITLE, "Enter valid number here")
input2 = InputBox("Input two: ", BOX_TITLE, "Enter valid number here")
intInputOne = CInt(input1)
intInputTwo = CInt(input2)

result = Addition(intInputOne, intInputTwo)
Call DisplayMsg ("The result is", result)

result = Subtraction(intInputOne, intInputTwo)
Call DisplayMsg ("The result is", result)

result = Multiplication(intInputOne, intInputTwo)
Call DisplayMsg ("The result is", result)

result = Division(intInputOne, intInputTwo)
Call DisplayMsg ("The result is", result)