'------------------------------------------------------------------------------'
' Summary: Get the standard prompt for a method error message box.
' Parameter(s):
'   err - The object that contains the error information
' Return(s): The standard prompt for a method error message box.
' Date Created: 2026-04-19
' Date Last Modified: 2026-05-04
'------------------------------------------------------------------------------'
Public Function GetErrorMsgBoxPrompt(err As ErrObject)
  GetErrorMsgBoxPrompt = _
    "Error: " & err.Number & vbCrLf & err.Description
End Function
