'------------------------------------------------------------------------------'
' Summary: Creates a string representation of the specified member of the
'   ValidationSrcType enumeration.
' Parameter(s):
'   vst - A member of the ValidationSrcType enumeration.
' Returns - A string representation of the specified member of the
'   ValidationSrcType enumeration.
' Date Created: 2026-07-04
' Date Last Modified: 22026-07-12
'------------------------------------------------------------------------------'
Public Function ValSrcTypeToString(vst As ValidationSrcType) As String
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ValSrcTypeToString"
  
  Select Case vst
    Case vstNoValidation
      ValSrcTypeToString = "No Validation"
    Case vstHardCodedValue
      ValSrcTypeToString = "Hard-Coded"
    Case vstCellRef
      ValSrcTypeToString = "Cell Ref"
    Case vstDefinedName
      ValSrcTypeToString = "Defined Name"
    Case vstFunction
      ValSrcTypeToString = "Function"
    Case Else
      ValSrcTypeToString = "Undefined"
  End Select

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function