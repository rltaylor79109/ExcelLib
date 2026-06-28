'------------------------------------------------------------------------------'
' Purpose: Unprotects all of the sheets and charts in the Workbook.
' Parameter(s)
'   silentMode - If True, the method does not write messages to Debug.Print
'     or display a message to the user when complete. It is an optional
'     parameter with a default value of False.
' Date Created: 2026-03-12
' Date Last Modified: 2026-06-10
'------------------------------------------------------------------------------'
Public Sub UnprotectAllSheetsAndCharts(Optional silentMode As Boolean = False)
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "UnprotectAllSheetsAndCharts"
  
  Dim cht As Chart
  Dim curActiveSheet As Object
  Dim wb As Workbook
  Dim ws As Worksheet
  
  Set wb = ThisWorkbook
  Set curActiveSheet = wb.ActiveSheet
  
  For Each ws In wb.Worksheets
    If Not silentMode Then
      Debug.Print "Unprotecting " & ws.name
    End If
    ws.Unprotect
  Next
  
  For Each cht In wb.Charts
    If Not silentMode Then
      Debug.Print "Unprotecting " & cht.name
    End If
    cht.Unprotect
  Next
  
Exit_Proc:
  If Not curActiveSheet Is Nothing Then
    curActiveSheet.Activate
  End If

  If Not silentMode Then
    Const prompt As String = "Unprotection complete."
    Const buttons As Long = vbInformation ' compiler fails if defined as vbMsgBoxStyle
    Const title As String = "Unprotecting All Sheets and Charts"
    MsgBox prompt, buttons, title
  End If

  Exit Sub
Err_Proc:
    ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub