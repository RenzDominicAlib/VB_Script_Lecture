' ============================================================
'Using Built in function with strings, MID, LEN


' ============================================================

' Using ByRef and ByVal with parameters/arguments

' ByVal ==> Passed by value
' ByRef ==> Passed by value or reference
' =============================================================

' option explicit
' on error resume next

'Dim site 

site = "www.youtube.com"
message = "I am learning basic VBScripting"
message1 = "www"
message2 = "youtube"
message3 = "com"

mynum1 = 10 'integer
mynum2 = "10" 'string

Sub DisplayMsg (strMsg, intResult)
	MsgBox strMsg & " : " & intResult, 0, BOX_TITLE
End Sub

'mid is used to return specific number of characters from a string
'Syntax : mid(string, start_position ,length[optional])

result1 = mid(site, 5, 2)
Call DisplayMsg("example 1", result1)

result2 = mid(site, 5, 10)
Call DisplayMsg("example 2", result2)

result3 = mid(site, 5)
Call DisplayMsg("example 3", result3)

result4 = message1 & "." & message2 &"." & message3
Call DisplayMsg("example 4", result4)

'Reversing the string
result5 = StrReverse(message2)
Call DisplayMsg("example 5", result5)

'Length of the string
result6 = Len(message2)
Call DisplayMsg("example 6", result6)