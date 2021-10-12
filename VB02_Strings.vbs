' ============================================================
				
' 				Header, comments, dates, important notes

' 1. ( vbLf ) - is used to declare a newline.
' 2. ( &_ ) - is use to break line in a code editor 
' 3. ( option explicit ) - used to strict the code in declaring variables
' 4. ( on error resume next ) - used to continue the running/execution disregarding the error.
' 5. Ucase(str) - convert to uppercase
' 6. Lcase(str) - convert to uppercase
' 7. Len(str) - convert to uppercase
' 8. replace(str) - convert to uppercase
'=============================================================

' ============================================================

' Variables, Constants, MsgBox, Commands, Line Continuation option
' explicit, On error Resume Next
'
' ============================================================

option explicit
on error resume next

'declaration of variables and constant
dim str
dim str2

str = "This is a string message"
str2 = "continuation message"

' MsgBox str
' MsgBox ucase(str)
' MsgBox lcase(str)
' MsgBox len(str)
' MsgBox replace(str, "i", "!", 1, 2, 1)
' MsgBox str & space(10) & str2
' MsgBox str 
' MsgBox Left(str, 5)
' MsgBox str 
' MsgBox Right(str, 5)
' MsgBox str 
' MsgBox Mid(str, 5, 9)
' MsgBox str 
' MsgBox len(trim(str))
' MsgBox str 
' MsgBox len(ltrim(str))
' MsgBox str 
' MsgBox len(rtrim(str))
' MsgBox str 
' MsgBox strreverse(str)
MsgBox str 
MsgBox strcomp(str, "string", 1)
