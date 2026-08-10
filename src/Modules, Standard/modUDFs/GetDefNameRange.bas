'------------------------------------------------------------------------------'
' Summary: Safely retrieves the Range object that a global defined name refers
'   to.
' Remarks: Checks if the defined name refers to a valid Range vs. a constant,
'   formula, or broken reference.
' Parameter(s):
'   pDefNameStr - A string that represents the defined name.
' Return(s): If the defined name refers to a cell range, the range object
'   that a global defined name refers to. Otherwise, Nothing if it refers to
'   a constant/formula, invalid reference, or if the defined name does not
'   exist.
' Date Created: 2026-08-08
' Date Last Modified: 2026-08-08
'------------------------------------------------------------------------------'
Public Function GetDefNameRange(ByVal pDefNameStr As String) As Range
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "GetDefNameRange"
    
  Dim nm As Name
  Dim targetRng As Range
  
  ' Try to fetch the defined name
  On Error Resume Next
  Set nm = ThisWorkbook.Names(pDefNameStr)
  err.Clear
  On Error GoTo Err_Proc
  
  ' If the defined name doesn't exist, return Nothing
  If nm Is Nothing Then
    Set GetDefNameRange = Nothing
    GoTo Exit_Proc
  End If
  
  ' Check if the defined name resolves to a Range object.
  ' Note: .RefersToRange throws an error (1004) if the name refers to
  ' a constant, non-range formula, or broken reference (#REF!).
  On Error Resume Next
  Set targetRng = nm.RefersToRange
  err.Clear
  On Error GoTo Err_Proc
  
  ' If targetRng is valid, return the Range; otherwise return Nothing
  If Not targetRng Is Nothing Then
    Set GetDefNameRange = targetRng
  Else
    Set GetDefNameRange = Nothing
  End If
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function
