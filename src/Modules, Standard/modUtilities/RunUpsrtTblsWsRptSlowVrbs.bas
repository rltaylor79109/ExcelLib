'------------------------------------------------------------------------------'
' Summary: Runs the UpsertDefNamesWsRefRpt with the fastMode parameter set to
'   false so that the table references and references count are shown
'   in the report. This is dramatically slower.  The silent parameter is
'   set to False so that the usual update complete message is displayed to
'   the user.
' Remarks: It is a wrapper that allows this method to be assigned as the macro
'   for a button on a worksheet.
' Date Created: 2026-07-17
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Public Sub RunUpsrtTblsWsRptSlowVrbs()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunUpsrtTblsWsRptSlowVrbs"

  UpsertTblsWsRefRpt fastMode:=False, silent:=False
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub