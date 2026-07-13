'------------------------------------------------------------------------------'
' Summary: Unprotects all of the sheets and charts in the Workbook.
' Parameter(s)
'   silentMode - If True, the method does not write messages to Debug.Print
'     or display a message to the user when complete. It is an optional
'     parameter with a default value of False.
' Date Created: 2026-03-12
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Public Sub UnprotectAllSheetsAndCharts(Optional silentMode As Boolean = False)
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "UnprotectAllSheetsAndCharts"

  Const title As String = "Unprotecting All Sheets and Charts"
  
  Dim buttons As VbMsgBoxStyle
  Dim cht As Chart
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
  
  For Each ws In wb.Worksheets
    If Not silentMode Then
      Debug.Print "Unprotecting " & ws.Name
    End If
    ws.Unprotect
  Next
  
  For Each cht In wb.Charts
    If Not silentMode Then
      Debug.Print "Unprotecting " & cht.Name
    End If
    cht.Unprotect
  Next
  
Exit_Proc:
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
  
  If Not curActiveSheet Is Nothing Then
    curActiveSheet.Activate
  End If
  
  If errorOccurred Then
    ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
    buttons = vbExclamation
    prompt = "The was a problem with the unprotection."
  Else
    buttons = vbExclamation
    prompt = "Unprotection complete."
  End If

  If Not silentMode Then
    MsgBox prompt, buttons, title
  End If

  Exit Sub
Err_Proc:
  errorOccurred = True
  Resume Exit_Proc
End Sub