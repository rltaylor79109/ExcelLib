'------------------------------------------------------------------------------'
' Summary: Runs the UpsertDataValAuditRptVerbose with the silent parameter
'   set to True. Other than error messages, no user messages are displayed
'   upon completion of the update.
' Remarks: It is a wrapper that allows this method to be assigned as the macro
'   for a button on a worksheet.
' Date Created: 2026-06-17
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Public Sub RunUpsrtDataValRptSilent()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunUpsrtDataValRptSilent"

  UpsertDataValAuditRpt silent:=True
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub