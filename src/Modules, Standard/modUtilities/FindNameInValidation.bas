'------------------------------------------------------------------------------'
' Summary: Finds a defined name in the data validation "Refers To" field
' Parameter(s)
'   searchName - The string to search for.
' Date Created: 2026-03-27
' Date Last Modified: 2026-03-27
'------------------------------------------------------------------------------'
Public Sub FindNameInValidation(searchName As String)
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "FindNameInValidation"
  
  Dim wrksht As Worksheet
  Dim theCell As Range
  For Each wrksht In ThisWorkbook.Worksheets
    On Error Resume Next
    For Each theCell In wrksht.Cells.SpecialCells(xlCellTypeAllValidation)
      If InStr(1, theCell.Validation.Formula1, searchName) > 0 Then
        Debug.Print "Found " & searchName & " in " & wrksht.Name & "!" & theCell.Address
      End If
    Next theCell
    On Error GoTo 0
  Next wrksht
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, mModuleName, METHOD_NAME
  Resume Exit_Proc
End Sub