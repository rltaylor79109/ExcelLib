'------------------------------------------------------------------------------'
' Purpose:  Prints a list of the names of all the embedded charts in this
'   workbook to debug output.
' Date Created: 2026-05-17
' Date Last Modified: 2026-06-09
'------------------------------------------------------------------------------'
Sub ListAllEmbeddedChartNames()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ListAllEmbeddedChartNames"

  Dim chtObj As ChartObject
  Dim chtTitle As String
  Dim chtCount As Long
  Dim output As String
  Dim wb As Workbook
  Dim ws As Worksheet
  
  Set wb = ThisWorkbook
  chtCount = 0
  For Each ws In wb.Worksheets
    For Each chtObj In ws.ChartObjects
      chtCount = chtCount + 1
      If chtObj.Chart.HasTitle Then
        chtTitle = chtObj.Chart.ChartTitle.text
      Else
        chtTitle = "[No Title]"
      End If
      
      output = ws.name & ", " & chtObj.name & ", " & chtTitle
      Debug.Print output
    Next chtObj
  Next ws
  
  If chtCount = 0 Then
    Debug.Print "This workbook does not contain any embedded charts."
  End If
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub