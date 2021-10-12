' ============================================================
				
' 				Header, comments, dates, important notes

' 1. ( vbLf ) - is used to declare a newline.
' 2. ( &_ ) - is use to break line in a code editor .
' 3. ( option explicit ) - used to strict the code in declaring variables.
' 4. ( on error resume next ) - used to continue the running/execution disregarding the error.
' 5. ( ReDim ) - used to modify the array.
' 6. ( Preserve ) - used to preserve/retain the original elements of array.

' ============================================================

'Arrays Redim

' ============================================================

option explicit
' on error resume next

'declaration of variables and constant

Const BOX_TITLE = "FRIENDS"

'declaration of array
Dim arrFriends()

'processing => actual job
ReDim arrFriends(4)
arrFriends(0) = "Renz"
arrFriends(1) = "Dom"
arrFriends(2) = "Jon"
arrFriends(3) = "Faye"
arrFriends(4) = "Nick"
ReDim Preserve arrFriends(9)
arrFriends(8) = "Tom"

MsgBox arrFriends(4), 0, BOX_TITLE
MsgBox arrFriends(8), 0, BOX_TITLE