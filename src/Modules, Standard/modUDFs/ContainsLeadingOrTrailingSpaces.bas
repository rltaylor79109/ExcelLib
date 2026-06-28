'------------------------------------------------------------------------------'
' Purpose:  Determines if the specified string contains any leading or
'   trailing spaces.
' Arguments:
'   pText - The string to test.
' Return:   True if pText contains any leading spaces, trailing spaces, or
'   both; otherwise false.
' Date Created: 2026-05-18
' Date Last Modified: 2026-05-18
'------------------------------------------------------------------------------'
Public Function ContainsLeadingOrTrailingSpaces(pText As String) As Boolean
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ContainsLeadingOrTrailingSpaces"
  
  ContainsLeadingOrTrailingSpaces = pText <> Trim(pText)

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, mModuleName, METHOD_NAME
  Resume Exit_Proc
End Function