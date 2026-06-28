'------------------------------------------------------------------------------'
' Purpose:  Prints a list of the names of all the chart sheets in this
'   workbook to debug output.
' Date Created: 2026-05-17
' Date Last Modified: 2026-06-09
'------------------------------------------------------------------------------'
Sub ListAllChartSheetNames()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ListAllChartSheetNames"

  Dim chtTitle As String
  Dim cs As Chart
  Dim csCount As Long
  Dim wb As Workbook
  
  Set wb = ThisWorkbook
  csCount = 0
  For Each cs In wb.Charts
      csCount = csCount + 1
      If cs.HasTitle Then
        chtTitle = cs.ChartTitle.text
      Else
        chtTitle = "[No Title]"
      End If
      Debug.Print cs.name & ", " & chtTitle
  Next cs
  
  If csCount = 0 Then
    Debug.Print "This workbook does not contain chart sheets."
  End If
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub