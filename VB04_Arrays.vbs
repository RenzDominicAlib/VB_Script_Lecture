' ============================================================
				
' 				Header, comments, dates, important notes

' 1. ( vbLf ) - is used to declare a newline.
' 2. ( &_ ) - is use to break line in a code editor 
' 3. ( option explicit ) - used to strict the code in declaring variables
' 4. ( on error resume next ) - used to continue the running/execution disregarding the error.

' ============================================================

'Arrays

' ============================================================

option explicit
' on error resume next

'declaration of variables and constant
dim strName
dim intIDNumber
dim intScienceGrade
dim intAverage
const BOX_TITLE = "Student Score"

'declaration of array
dim arrGrades(4)

'processing => actual job

strName = "Renz Dominic"
intIDNumber = 1129989
arrGrades(0) = 95
arrGrades(1) = 96
arrGrades(2) = 97
arrGrades(3) = 98
arrGrades(4) = 99
intScienceGrade = arrGrades(0)
intAverage = (arrGrades(0) + arrGrades(1) + arrGrades(2) + arrGrades(3) + arrGrades(4)) / 5

MsgBox "Welcome " & strName & "!" & vbLf & "Student number: " & intIDNumber, 0, BOX_TITLE
MsgBox "Your Grades are:" & vbLf &_ 
 "Science: " & intScienceGrade & vbLf &_
 "Mathematics: " & arrGrades(1) & vbLf &_
 "Arts: " & arrGrades(2) & vbLf &_
 "English: " & arrGrades(3) & vbLf &_
 "Music: " & arrGrades(4), 0, BOX_TITLE

MsgBox "Your average score is: " & intAverage