'------------------------------------------------------------------------------'
' Summary: Gets a chartsheet object by its CodeName property.
' Remarks: A Chartsheet's CodeName property is distinct from its Name property.
'   The CodeName property is visible in the Project Explorer in the Excel
'   VBA IDE. The Name property is shown on the Chart tab in the regular
'   Excel window. The method searches the active workbook.
' Parameter(s):
'   pCodeName - Represents the value of the CodeName property of the chart of
'     interest.
' Return(s): The chartsheet object with the specified CodeName property.
' Date Created: 2026-05-17
' Date Last Modified: 2026-06-09
'------------------------------------------------------------------------------'
Public Function GetChartsheetByCodeName(pCodeName As String) As Chart
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "GetChartsheetByCodeName"
  
  Dim errorOccurred As Boolean
  Dim cs As Chart
  Dim wb As Workbook
  
  Application.EnableEvents = False
  errorOccurred = False
  Set wb = ThisWorkbook
  For Each cs In wb.Charts
    If UCase(cs.CodeName) = UCase(pCodeName) Then
      Set GetChartsheetByCodeName = cs
      GoTo Exit_Proc
    End If
  Next cs
  
Exit_Proc:
  Application.EnableEvents = True
  If errorOccurred Then
    ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  End If
  Exit Function
Err_Proc:
  errorOccurred = True
  Resume Exit_Proc
End Function