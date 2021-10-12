' ============================================================
				
' 				Header, comments, dates, important notes

' 1. ( vbLf ) - is used to declare a newline.
' 2. ( &_ ) - is use to break line in a code editor .
' 3. ( option explicit ) - used to strict the code in declaring variables.
' 4. ( on error resume next ) - used to continue the running/execution disregarding the error.
' 5. ( ReDim ) - used to modify the array.
' 6. ( Preserve ) - used to preserve/retain the original elements of array.

' ============================================================

'Arrays Declaration

' ============================================================

option explicit
' on error resume next

'declaration of variables and constant

Const BOX_TITLE = "FRIENDS"

'declaration of array
Dim arrFriends

'declaration of array elements
arrFriends = Array("Renz", "Jon", "Dom", "Faye", "Nick")

'adding and preserving array elements
ReDim Preserve arrFriends(9)
arrFriends(8) = "Marg"
arrFriends(9) = "June"

MsgBox arrFriends(0), 0, BOX_TITLE
MsgBox arrFriends(1), 0, BOX_TITLE
MsgBox arrFriends(2), 0, BOX_TITLE
MsgBox arrFriends(3), 0, BOX_TITLE
MsgBox arrFriends(4), 0, BOX_TITLE

MsgBox arrFriends(8), 0, BOX_TITLE
MsgBox arrFriends(9), 0, BOX_TITLE