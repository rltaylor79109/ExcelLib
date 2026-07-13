'------------------------------------------------------------------------------'
' Summary: Creates a string representation of the specified member of the
'   VbCallType enumeration.
' Parameter(s):
'   pCallType - A member of the vbCallType enumeration.
' Return(s): A string representation of the specified member of the
'   VbCallType enumeration.
' Date Created: 2026-05-08
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Function VbCallTypeToString(pCallType As VbCallType) As String
  On Error GoTo Err_Proc
  Const methodName As String = "VbCallTypeToString"
  
  Select Case pCallType
    Case VbGet
      VbCallTypeToString = "Get"
    Case VbLet
      VbCallTypeToString = "Let"
    Case VbMethod
      VbCallTypeToString = "Method"
    Case VbSet
      VbCallTypeToString = "Set"
    Case Else
      VbCallTypeToString = "Unknown"
  End Select

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, methodName
  Resume Exit_Proc
End Function