'------------------------------------------------------------------------------'
' Summary: Determines whether a string contains any of the characters that
'   are not permitted in sheet names. The prohibited characters are
'   \ / * [ ] ? :
' Parameter(s):
'   pSheetName - The string to test.
' Return:   True if pSheetName contains any of the prohibited characters,
'   otherwise false.
' Date Created: 2018-08-30
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Function ContainsIllegalWorksheetNameCharacters( _
  pSheetName As String) As Boolean
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ContainsIllegalWorksheetNameCharacters"
  
  ContainsIllegalWorksheetNameCharacters = _
    ContainsAnyChar(pText:=pSheetName, _
    charsToFind:=WORKSHEET_NAME_FORBIDDEN_CHARS)

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function