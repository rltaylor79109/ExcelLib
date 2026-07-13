'------------------------------------------------------------------------------'
' Summary: Runs the ProtectAllSheetsAndCharts using the default parameters.
' Remarks: It is a wrapper that allows this method to be assigned as the macro
'   for a button on a worksheet
' Date Created: 2026-05-08
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Sub RunProtectAllSheetsAndCharts()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunProtectAllSheetsAndCharts"

  ProtectAllSheetsAndCharts

Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub