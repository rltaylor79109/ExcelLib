'------------------------------------------------------------------------------'
' Summary: Gets the sum of cell numeric values in a range whose column is NOT
'   hidden and the cell does not contain an error.
' Parameter(s):
'   rng - The range to sum.
' Return(s): The sum of cell numeric values in a range whose column is NOT
'   hidden and the cell does not contain an error.
' Date Created: 2026-08-19
' Date Last Modified: 2026-08-19
'------------------------------------------------------------------------------'
Public Function SUM_VISIBLE_ROW(rng As Range) As Double
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "SUM_VISIBLE_ROW"
  
  Dim cell As Range
  Dim total As Double
    
  For Each cell In rng.Cells
    ' Only sum cells whose column is NOT hidden and does NOT contain an error
    If _
      Not cell.EntireColumn.Hidden _
      And _
      Not IsError(cell.value) _
      And _
      IsNumeric(cell.value) _
      Then
        total = total + cell.value
    End If
  Next cell
    
  SUM_VISIBLE_ROW = total
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function