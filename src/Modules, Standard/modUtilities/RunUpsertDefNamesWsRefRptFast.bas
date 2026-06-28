'------------------------------------------------------------------------------'
' Summary: Runs the ProtectAllSheetsAndCharts with the fast parameter set to
'   true. This is dramatically faster but the defined names references and
'   references counts are not shown in the report.
' Remarks: It is a wrapper that allows this method to be assigned as the macro
'   for a button on a worksheet.
' Date Created: 2026-05-13
' Date Last Modified: 2026-06-14
'------------------------------------------------------------------------------'
Public Sub RunUpsertDefNamesWsRefRptFast()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunUpsertDefNamesWsRefRptFast"

  UpsertDefNamesWsRefRpt fast:=True
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub