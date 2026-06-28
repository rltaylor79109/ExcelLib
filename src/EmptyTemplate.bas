Option Explicit

'------------------------------------------------------------------------------'
' Class Name: 
' Summary: 
' Date Created: 2026-05-04
' Date Last Modified: 2026-05-04
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Fields
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Constant Fields
'------------------------------------------------------------------------------'

' The name of this module.
Private Const mModuleName As String = "modModuleName"

'------------------------------------------------------------------------------'
' Variable Fields
'------------------------------------------------------------------------------'

'
Private mMyField as Integer

'------------------------------------------------------------------------------'
' Properties
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Summary: 
' Value:
' Date Created: 2026-05-04
' Date Last Modified: 2026-05-04
'------------------------------------------------------------------------------'
Property Get myProperty() As Boolean
  myProperty = True
End Property
Property Let myProperty(value As Boolean)
  myProperty = value
End Property

'------------------------------------------------------------------------------'
' Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Public Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Summary: Template for Public method.
' Parameter(s)
'   paramName - 
' Date Created: 2026-05-04
' Date Last Modified: 2026-05-04
'------------------------------------------------------------------------------'
Public Sub PublicTemplate(paramName As String)
  On Error GoTo Err_Proc
  Const methodName As String = "PublicTemplate"

Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, mModuleName, methodName
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Get the standard prompt for a method error message box.
' Parameter(s):
'   err - The object that contains the error information
' Return(s): the standard prompt for a method error message box.
' Date Created: 2026-04-19
' Date Last Modified: 2026-04-19
'------------------------------------------------------------------------------'
Public Function GetErrorMsgBoxPrompt(err As ErrObject)
  GetErrorMsgBoxPrompt = _
    "Error: " & err.Number & vbCrLf & err.Description
End Function

End Sub

'------------------------------------------------------------------------------'
' Summary: Shows the standard method error message box.
' Parameter(s):
'   err - The object that contains the error information.
'   moduleName - The name of the module that contains the method that produced
'     the error.
'   methodName - The name of the method that produced the error.
' Date Created: 2026-04-19
' Date Last Modified: 2026-04-19
'------------------------------------------------------------------------------'
Public Sub ShowMethodErrorMsgBox( _
  err As ErrObject, _
  moduleName As String, _
  methodName As String)

  Dim prompt As String
  Dim title As String
  
  prompt = "Source: Method " & moduleName & "." & methodName & vbCrLf & _
    "Error " & err.Number & ": " & err.Description
  title = moduleName & "." & methodName
  MsgBox prompt, vbExclamation, title
End Sub

'------------------------------------------------------------------------------'
' Summary: Shows the standard property error message box.
' Parameter(s):
'   err - The object that contains the error information.
'   moduleName - The name of the module that contains the property that produced
'     the error.
'   propertyName - The name of the property that produced the error.
'   callType - The property call type.
' Date Created: 2026-04-19
' Date Last Modified: 2026-04-19
'------------------------------------------------------------------------------'
Public Sub ShowPropertyErrorMsgBox( _
  err As ErrObject, _
  moduleName As String, _
  propertyName As String, _
  callType As VbCallType)
  
  On Error GoTo Err_Proc
  Const methodName As String = "ShowPropertyErrorMsgBox"
  
  Dim callTypeString As String
  Dim prompt As String
  Dim title As String
  
  Select Case callType
    Case VbGet
      callTypeString = "Get"
    Case VbLet
      callTypeString = "Let"
    Case VbMethod
      callTypeString = "Method"
    Case VbSet
      callTypeString = "Set"
  End Select
  
  prompt = "Source: Property " & moduleName & "." & propertyName & "." & _
    callTypeString & vbCrLf & _
    "Error " & err.Number & ": " & err.Description
  title = moduleName & "." & propertyName & "." & callTypeString
  MsgBox prompt, vbExclamation, title
    
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, mModuleName, methodName
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Private Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Summary: Template for Private method.
' Parameter(s)
'   paramName - 
' Date Created: 2026-05-04
' Date Last Modified: 2026-05-04
'------------------------------------------------------------------------------'
Private Sub PrivateTemplate(paramName As String)
  On Error GoTo Err_Proc
  Const methodName As String = "PrivateTemplate"

Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, mModuleName, methodName
  Resume Exit_Proc
End Sub