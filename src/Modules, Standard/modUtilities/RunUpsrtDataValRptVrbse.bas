'------------------------------------------------------------------------------'
' Summary: Runs the UpsertDataValAuditRptVerbose with the silent parameter
'   set to False. An update complete message is displayed to the user.
' Remarks: It is a wrapper that allows this method to be assigned as the macro
'   for a button on a worksheet.
' Date Created: 2026-06-17
' Date Last Modified: 2026-06-17
'------------------------------------------------------------------------------'
Public Sub RunUpsrtDataValRptVrbse()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunUpsrtDataValRptVrbse"

  UpsertDataValAuditRpt silent:=False
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub