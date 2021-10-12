' ============================================================
				
' 				Header, comments, dates, important notes

' 1. ( vbLf ) - is used to declare a newline.
' 2. ( &_ ) - is use to break line in a code editor 
' 3. ( option explicit ) - used to strict the code in declaring variables
' 4. ( on error resume next ) - used to continue the running/execution disregarding the error.


' ============================================================

' Variables, Constants, MsgBox, Commands, Line Continuation option
' explicit, On error Resume Next
'
' ============================================================

option explicit
on error resume next

'declaration of variables and constant
dim intScienceGrade
dim intMathGrade
dim intAverage
dim strName
dim intIDNumber
const BOX_TITLE = "Student Score"

'processing => actual job

strName = "Renz Alib"
intIDNumber = 1129989
intScienceGrade = 95
intMathGrade = 75
intAverage = (intScienceGrade + intMathGrade) / 2

MsgBox "Name: " &strName & vbLf & "ID Number: " & intIDNumber, 0, BOX_TITLE

MsgBox "You've got " & vbLf &"Science: " & intScienceGrade & vbLf &_
 "Mathematics: " & intMathGrade, 0, BOX_TITLE

MsgBox "Your average score is: " & intAverage, 0, BOX_TITLE