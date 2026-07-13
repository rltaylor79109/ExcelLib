'------------------------------------------------------------------------------'
' Summary: Determines if the specified string contains any leading or
'   trailing spaces.
' Parameter(s):
'   pText - The string to test.
' Return:   True if pText contains any leading spaces, trailing spaces, or
'   both; otherwise false.
' Date Created: 2026-05-18
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Function ContainsLeadingOrTrailingSpaces(pText As String) As Boolean
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ContainsLeadingOrTrailingSpaces"
  
  ContainsLeadingOrTrailingSpaces = pText <> Trim(pText)

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function