' ============================================================
'Using Built in function


' ============================================================

'IsNull, IsEmpty, Null, Empty

' =============================================================

' option explicit
' on error resume next

Function DisplayMsg (strMsg, intResult)
	MsgBox strMsg & " : " & intResult, 0, BOX_TITLE
End Function

Dim var1, var2, var3, var4, var5

result1 = IsNull(var1)
DisplayMsg "Example 1", result1

var2 = "Iam Here!"
result2 = IsNull(var2)
DisplayMsg "Example 2", result2

va3 = Null
resul3 = IsNull(va3)
DisplayMsg "Example 3", resul3

var4 = Empty
result4 = IsNull(var4)
DisplayMsg "Example 4", result4

' var5 = Null
var5 = "xyz" 
condition = IsNull(var5)

If condition = True Then
	result5 = "I am NULL"
Else
	result5 = "I am NOT NULL"
End If

DisplayMsg "Example 5", result5