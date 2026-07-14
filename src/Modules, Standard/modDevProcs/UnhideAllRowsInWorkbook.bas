'------------------------------------------------------------------------------'
' Summary: Unhides all Worksheets and Chartsheets in this Workbook.
' Date Created: 2026-07-14
' Date Last Modified: 2026-07-14
'------------------------------------------------------------------------------'
Sub UnhideAllRowsInWorkbook()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "UnhideAllRowsInWorkbook"
  
  Const title As String = "Unhiding All Rows in Worksheets"
  
  Dim buttons As VbMsgBoxStyle
  Dim curActiveSheet As Object
  Dim errorOccurred As Boolean
  Dim prompt As String
  Dim wb As Workbook
  Dim ws As Worksheet
    
  errorOccurred = False
  
  ' Optimization: Turn off UI updates
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False
    
  Set wb = ThisWorkbook
  Set curActiveSheet = wb.ActiveSheet
  
  ' Loop through each worksheet in the active workbook
  For Each ws In wb.Worksheets
    ws.Rows.Hidden = False
  Next ws

exit_Proc:
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
  
  If Not curActiveSheet Is Nothing Then
    curActiveSheet.Activate
  End If
  
  If errorOccurred Then
    ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
    buttons = vbExclamation
    prompt = "The was a problem unhidding all the rows."
  Else
    buttons = vbInformation
    prompt = "All Rows in all Worksheets have been unhidden."
  End If
  
  MsgBox prompt, buttons, title

  Exit Sub
Err_Proc:
  errorOccurred = True
  Resume exit_Proc
End Sub