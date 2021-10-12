' ============================================================
'Using Built in function Formatting


' ============================================================

'Numbers
'FormatNumber(Expression, NumdigitAfterDecimal, IncludeLeadingDigit, UseParenthesisForNegativeNumbers, GroupDigits)

' =============================================================

' option explicit
' on error resume next

Function DisplayMsg (strMsg, intResult)
	MsgBox strMsg & " : " & intResult, 0, BOX_TITLE
End Function

Dim number
number = -12345.6789123

DisplayMsg "Example 1", number
DisplayMsg "Example 2", FormatNumber (number, 2)
DisplayMsg "Example 3", FormatNumber (number, 2, ,vbTrue)
DisplayMsg "Example 4", FormatNumber (number, 2, ,vbTrue, vbFalse)

' =============================================================
'Percentage

Dim MyPercent
MyPercent = FormatPercent(25/50)

DisplayMsg "Example 5", MyPercent