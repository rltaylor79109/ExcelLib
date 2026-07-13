'------------------------------------------------------------------------------'
' Summary: Creates a string representation of the specified member of the
'   XlDVType enumeration (Validation.Type).
' Parameter(s):
'   dvType - A member of the XlDVType enumeration.
' Return(s): A string representation of the specified member of the
'   XlDVType enumeration (Validation.Type).
' Date Created: 2026-07-07
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Function XlDVTypeToString(dvType As XlDVType) As String
  On Error GoTo Err_Proc
  Const methodName As String = "XlDVTypeToString"
  
  Select Case dvType
    Case xlValidateInputOnly
      XlDVTypeToString = "Input Only"
    Case xlValidateWholeNumber
      XlDVTypeToString = "Whole Number"
    Case xlValidateDecimal
      XlDVTypeToString = "Decimal"
    Case xlValidateList
      XlDVTypeToString = "List"
    Case xlValidateDate
      XlDVTypeToString = "Date"
    Case xlValidateTime
      XlDVTypeToString = "Time"
    Case xlValidateTextLength
      XlDVTypeToString = "Text Length"
    Case xlValidateCustom
      XlDVTypeToString = "Custom"
    Case Else
      XlDVTypeToString = "Undefined"
  End Select

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, methodName
  Resume Exit_Proc
End Function