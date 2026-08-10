'------------------------------------------------------------------------------'
' Summary: Safely retrieves the value of any global defined name
' (whether it refers to a Range or a direct Constant/Formula).
' Remarks: This avoids errors from active sheet changes or duplicate locally
'   scoped defined names.
' Parameter(s):
'   pDefNameStrStr - A string that represents the the defined name.
' Return(s): The worksheet object with the specifed CodeName property.
' Date Created: 2026-07-22
' Date Last Modified: 2026-07-22
'------------------------------------------------------------------------------'
Public Function GetDefNameVal(ByVal pDefNameStr As String) As Variant
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "GetDefinedNameValue"
    
  Dim nm As Name
  Dim targetRng As Range
  
  On Error Resume Next
  Set nm = ThisWorkbook.Names(pDefNameStr)
  err.Clear
  On Error GoTo Err_Proc
  
  ' If the name doesn't exist, return Empty/Error
  If nm Is Nothing Then
    GetDefNameVal = Empty
    GoTo Exit_Proc
  End If
  
  ' Check if the name actually points to a worksheet Range
  On Error Resume Next
  Set targetRng = nm.RefersToRange
  err.Clear
  On Error GoTo Err_Proc
  
  If Not targetRng Is Nothing Then
    ' It's a range reference (e.g., refers to $C$1)
    GetDefNameVal = targetRng.value
  Else
    ' It's a constant or formula (e.g., refers to =7)
    GetDefNameVal = Application.Evaluate(nm.RefersTo)
  End If
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function