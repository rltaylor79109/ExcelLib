'------------------------------------------------------------------------------'
' Summary: Shows the standard method error message box.
' Parameter(s):
'   err - The object that contains the error information.
'   moduleName - The name of the module that contains the method that produced
'     the error.
'   METHOD_NAME - The name of the method that produced the error.
' Date Created: 2026-04-19
' Date Last Modified: 2026-05-08
'------------------------------------------------------------------------------'
Public Sub ShowMethodErrorMsgBox( _
  err As ErrObject, _
  moduleName As String, _
  METHOD_NAME As String)

  Dim prompt As String
  Dim buttons As VbMsgBoxStyle
  Dim title As String
  
  prompt = "Source: Method " & moduleName & "." & METHOD_NAME & vbCrLf & _
    "Error " & err.Number & ": " & err.Description
  buttons = VbMsgBoxStyle.vbExclamation
  title = moduleName & "." & METHOD_NAME
  MsgBox prompt, buttons, title
End Sub