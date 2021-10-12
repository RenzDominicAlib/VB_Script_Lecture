' ============================================================
				
' 				Header, comments, dates, important notes

' 1. ( vbLf ) - is used to declare a newline.
' 2. ( &_ ) - is use to break line in a code editor 
' 3. ( option explicit ) - used to strict the code in declaring variables
' 4. ( on error resume next ) - used to continue the running/execution disregarding the error.

' ============================================================

'Arrays 2D

' ============================================================

option explicit
' on error resume next

'declaration of variables and constant
dim strClassSection
dim strGradeLevel
const BOX_TITLE = "Student Favorite"

'declaration of array
dim arrFavSubject(1,4)

'processing => actual job

strClassSection = "Venus"
strGradeLevel = "Grade 6"

arrFavSubject(0,0) = "Renz"
arrFavSubject(0,1) = "Dom"
arrFavSubject(0,2) = "Jon"
arrFavSubject(0,3) = "Faye"
arrFavSubject(0,4) = "Nick"

arrFavSubject(1,0) = "Mathematics"
arrFavSubject(1,1) = "Science"
arrFavSubject(1,2) = "Music"
arrFavSubject(1,3) = "Arts"
arrFavSubject(1,4) = "English"


MsgBox "Welcome " & strGradeLevel & " Section " & strClassSection, 0, BOX_TITLE
MsgBox "Here are the outstanding students: " & vbLf &_ 
 "1. " & arrFavSubject(0,0) & " in " & arrFavSubject(1,0) & vbLf &_
 "2. " & arrFavSubject(0,1) & " in " & arrFavSubject(1,1) & vbLf &_
 "3. " & arrFavSubject(0,2) & " in " & arrFavSubject(1,2) & vbLf &_
 "4. " & arrFavSubject(0,3) & " in " & arrFavSubject(1,3) & vbLf &_
 "5. " & arrFavSubject(0,4) & " in " & arrFavSubject(1,4), 0, BOX_TITLE
