'------------------------------------------------------------------------------'
' Summary: Unhides all Worksheets and Chartsheets in this Workbook.
' Date Created: 2017-03-24
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Sub UnhideAllSheetsAndCharts()
  On Error GoTo Err_Proc
  Const methodName As String = "UnhideAllSheetsAndCharts"
  
  Dim cht As Chart
  Dim wb As Workbook
  Dim ws As Worksheet
  
  Set wb = ThisWorkbook
  For Each ws In wb.Worksheets
    ws.Visible = xlSheetVisible
  Next ws
  
  For Each cht In wb.Charts
    cht.Visible = xlSheetVisible
  Next cht

Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, methodName
  Resume Exit_Proc
End Sub