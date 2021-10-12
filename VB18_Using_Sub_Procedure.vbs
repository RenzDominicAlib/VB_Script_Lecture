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

MsgBox "M1 : Hello!!! How are you?", 0, BOX_TITLE
Call DisplayMessages
MsgBox "M2 : Hello!!! How are you?", 0, BOX_TITLE
MsgBox "M3 : Hello!!! How are you?", 0, BOX_TITLE
Call DisplayMessages

Sub DisplayMessages
	MsgBox "1 Im a sub message", 0, BOX_TITLE
	MsgBox "2 Im a sub message", 0, BOX_TITLE
	MsgBox "3 Im a sub message", 0, BOX_TITLE
End Sub