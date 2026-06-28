'------------------------------------------------------------------------------'
' Purpose:  Escapes characters that have special meanings in RegEx in the
'   specified string.
' Parameter(s):
'   pString - The string to modify.
' Returns: The input string modified so that characters that have special
'   meanings in RegEx are escaped.
' Date Created: 2026-06-15
' Date Last Modified: 2026-06-15
'------------------------------------------------------------------------------'
Public Function EscapeRegexChars(ByVal pString As String) As String
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "EscapeRegexChars"
  
  Const SPECIAL_CHARACTERS As String = "\^$.|?*+()[]{}"
  
  Dim char As String
  Dim escapedStr As String
  Dim i As Long

  escapedStr = pString
  For i = 1 To Len(SPECIAL_CHARACTERS)
    char = Mid(SPECIAL_CHARACTERS, i, 1)
    escapedStr = Replace(escapedStr, char, "\" & char)
  Next i
  EscapeRegexChars = escapedStr
    
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function