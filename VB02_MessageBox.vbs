'================================================================================================================================

'MsgBox(prompt[,buttons][,title][,helpfile,context])

' 0 = vbOKOnly - OK button only
' 1 = vbOKCancel - OK and Cancel buttons
' 2 = vbAbortRetryIgnore - Abort, Retry, and Ignore buttons
' 3 = vbYesNoCancel - Yes, No, and Cancel buttons
' 4 = vbYesNo - Yes and No buttons
' 5 = vbRetryCancel - Retry and Cancel buttons
' 16 = vbCritical - Critical Message icon
' 32 = vbQuestion - Warning Query icon
' 48 = vbExclamation - Warning Message icon
' 64 = vbInformation - Information Message icon
' 0 = vbDefaultButton1 - First button is default
' 256 = vbDefaultButton2 - Second button is default
' 512 = vbDefaultButton3 - Third button is default
' 768 = vbDefaultButton4 - Fourth button is default
' 0 = vbApplicationModal - Application modal (the current application will not work until the user responds to the message box)
' 4096 = vbSystemModal - System modal (all applications wont work until the user responds to the message box)


'================================================================================================================================

dim msg, boxtitle

msg = "Are you having fun in learning VBScript?"
boxtitle = "Box Title"

'correct
MsgBox msg, 1, boxtitle

'incorrect
MsgBox boxtitle, 0, msg

MsgBox"Hello everyone!",65,"Example"
