'------------------------------------------------------------------------------'
' Summary: Runs the UpsertDefNamesAllRefsRpt with the fastMode parameter set to
'   True so that the defined names worksheet references report is not updated
'   before the defined names all reference report is updated. This is
'   dramatically faster.
' Remarks: It is a wrapper that allows this method to be assigned as the macro
'   for a button on a worksheet.
' Date Created: 2026-06-17
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Sub RunUpsertDefNamesAllRptFast()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunUpsertDefNamesAllRptFast"

  UpsertDefNamesAllRefsRpt fastMode:=True
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub