'------------------------------------------------------------------------------'
' Summary: Determines if the specified cell has Data Validation.
' Parameter(s):
'   cell - Represents the target cell Range.
' Return(s): True if the specified cell has Data Validation; otherwise, False.
' Date Created: 2026-08-02
' Date Last Modified: 2026-08-02
'------------------------------------------------------------------------------'
Public Function HasValidation(cell As Range) As Boolean
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "HasValidation"
  
  On Error Resume Next
  HasValidation = (cell.Validation.Type >= 0)
  err.Clear
  On Error GoTo Err_Proc
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function