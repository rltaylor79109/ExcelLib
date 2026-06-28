'------------------------------------------------------------------------------'
' Summary: Shows the standard property error message box.
' Parameter(s):
'   err - The object that contains the error information.
'   moduleName - The name of the module that contains the property that produced
'     the error.
'   propertyName - The name of the property that produced the error.
'   callType - The property call type.
' Date Created: 2026-04-19
' Date Last Modified: 2026-05-08
'------------------------------------------------------------------------------'
Public Sub ShowPropertyErrorMsgBox( _
  err As ErrObject, _
  moduleName As String, _
  propertyName As String, _
  callType As VbCallType)
  
  Dim buttons As VbMsgBoxStyle
  Dim callTypeString As String
  Dim prompt As String
  Dim title As String
  
  callTypeString = VbCallTypeToString(callType)
  prompt = "Source: Property " & moduleName & "." & propertyName & "." & _
     callTypeString & vbCrLf & _
    "Error " & err.Number & ": " & err.Description
  buttons = VbMsgBoxStyle.vbExclamation
  title = moduleName & "." & propertyName & "." & callTypeString
  MsgBox prompt, buttons, title
End Sub