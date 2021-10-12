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
Dim temp, intTemp

Function convertFtoC(ftemp)
	MsgBox "Message box inside the sub Function"
	convertFtoC = (ftemp - 32) * (5/9)
End Function

Sub ConvertTemp
	temp = InputBox("Actual Temperature to convert: ", BOX_TITLE, "Enter valid Temperature in Fahrenheit here")
	intTemp = CInt(temp)
	MsgBox "Message box inside the sub Procedure"
	MsgBox "The converted value is: " & convertFtoC(intTemp), 0, BOX_TITLE
End Sub

MsgBox "Conversion code from Celsius to Fahrenheit"
Call ConvertTemp