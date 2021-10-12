' ============================================================
				
' 				Header, comments, dates, important notes

' 1. ( vbLf ) - is used to declare a newline.
' 2. ( &_ ) - is use to break line in a code editor .
' 3. ( option explicit ) - used to strict the code in declaring variables.
' 4. ( on error resume next ) - used to continue the running/execution disregarding the error.
' 5. ( InputBox() ) - used to create a inputbox.
' 6. ( CInt() ) - converting into the real integers


' ============================================================

' Sub Procedure --- Does not return a value
' Function Procedure --- can return a value
' Calculator code


' ============================================================

option explicit
' on error resume next

'*************************************************************
'sub to read external vbs files

Sub Include(extVBScriptFile)
	Dim objFso, objExtFile
	Dim strfileContent, strScriptDir

	strScriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(Wscript.ScriptFullName)

	Set objFso = CreateObject("Scripting.FileSystemObject")
	Set objExtFile = objFso.OpenTextFile(strScriptDir & "\" & extVBScriptFile)
	strfileContent = objExtFile.ReadAll
	ExecuteGlobal strfileContent
End Sub
'*************************************************************
'Should always above
Call Include ("VB22_Function.vbs")
'Include "VB22_Function.vbs"

Dim a, b, result
a = 10
b = 5

result = Addition(a, b)
DisplayMsg "The result is", result

result = Subtraction(a, b)
DisplayMsg "The result is", result

result = Multiplication(a, b)
Call DisplayMsg ("The result is", result)

result = Division(a, b)
Call DisplayMsg ("The result is", result)