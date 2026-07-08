'------------------------------------------------------------------------------'
' Summary: Creates a string representation of a member of the
'   XlFormatConditionOperator enumeration.
' Parameter(s):
'   op - A member of the XlFormatConditionOperator enumeration.
' Date Created: 2026-07-07
' Date Last Modified: 2026-07-07
'------------------------------------------------------------------------------'
Public Function XlFormatConditionOperatorToString( _
  op As XlFormatConditionOperator) As String
  
  On Error GoTo Err_Proc
  Const methodName As String = "XlFormatConditionOperatorToString"
  
  Select Case op
    Case xlBetween
      XlFormatConditionOperatorToString = "Between"
    Case xlNotBetween
      XlFormatConditionOperatorToString = "Not Between"
    Case xlEqual
      XlFormatConditionOperatorToString = "Equal"
    Case xlNotEqual
      XlFormatConditionOperatorToString = "Not Equal"
    Case xlGreater
      XlFormatConditionOperatorToString = "Greater"
    Case xlLess
      XlFormatConditionOperatorToString = "Less"
    Case xlGreaterEqual
      XlFormatConditionOperatorToString = "Greater Equal"
    Case xlLessEqual
      XlFormatConditionOperatorToString = "Less Equal"
    Case Else
      XlFormatConditionOperatorToString = "Undefined"
  End Select

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, methodName
  Resume Exit_Proc
End Function