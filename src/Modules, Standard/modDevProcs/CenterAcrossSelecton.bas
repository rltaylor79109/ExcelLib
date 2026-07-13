'------------------------------------------------------------------------------'
' Summary: Sets the horizontal aligment format of the current selection to
'   center across selection.
' Date Created: 2026-05-14
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Sub CenterAcrossSelecton()
  On Error GoTo Err_Proc
  Const methodName As String = "CenterAcrossSelecton"

  Selection.HorizontalAlignment = xlCenterAcrossSelection

Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, methodName
  Resume Exit_Proc
End Sub