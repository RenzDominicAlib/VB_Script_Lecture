' ============================================================
				
' 				Header, comments, dates, important notes

' 1. ( vbLf or VBCr) - is used to declare a newline.
' 2. ( &_ ) - is use to break line in a code editor .
' 3. ( option explicit ) - used to strict the code in declaring variables.
' 4. ( on error resume next ) - used to continue the running/execution disregarding the error.
' 5. ( ReDim ) - used to modify the array.
' 6. ( Preserve ) - used to preserve/retain the original elements of array.
' 7. ( Split ) - used to Split elements/string

' ============================================================

'Arrays Split Strings

' ============================================================

option explicit
' on error resume next

'declaration of variables and constant

Const WEBSITE = "www.google.com"

'declaration of array
Dim arrNewStrings

'spliting strings
arrNewStrings = Split(WEBSITE, ".")


MsgBox arrNewStrings(0), 0, WEBSITE
MsgBox arrNewStrings(1), 0, WEBSITE
MsgBox arrNewStrings(2), 0, WEBSITE
