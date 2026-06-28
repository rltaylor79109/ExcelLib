'------------------------------------------------------------------------------'
' Summary: Runs the ProtectAllSheetsAndCharts with the fast parameter set to
'   false so that the defined names references and references count are shown
'   in the report. This is dramatically slower.
' Remarks: It is a wrapper that allows this method to be assigned as the macro
'   for a button on a worksheet.
' Date Created: 2026-05-13
' Date Last Modified: 2026-06-14
'------------------------------------------------------------------------------'
Public Sub RunUpsertDefNamesWsRefRptSlow()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunUpsertDefNamesWsRefRptSlow"

  UpsertDefNamesWsRefRpt fast:=False
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub