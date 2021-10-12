' ============================================================
				
' 				Header, comments, dates, important notes

' 1. ( vbLf or VBCr) - is used to declare a newline.
' 2. ( &_ ) - is use to break line in a code editor .
' 3. ( option explicit ) - used to strict the code in declaring variables.
' 4. ( on error resume next ) - used to continue the running/execution disregarding the error.
' 5. ( ReDim ) - used to modify the array.
' 6. ( Preserve ) - used to preserve/retain the original elements of array.
' 7. ( Join ) - used to join string

' ============================================================

'Arrays Split Strings

' ============================================================

option explicit
' on error resume next

'declaration of variables and constant
Dim strNewString
Const WEBSITE = "www.google.com"

'declaration of array and variables

Dim arrSite(2)

arrSite(0) = "www"
arrSite(1) = "accenture"
arrSite(2) = "com"


'join strings
strNewString = Join(arrSite, ".")


MsgBox strNewString, 0, WEBSITE

