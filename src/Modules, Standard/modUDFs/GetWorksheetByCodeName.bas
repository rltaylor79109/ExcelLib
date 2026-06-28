'------------------------------------------------------------------------------'
' Summary: Gets a worksheet object by its CodeName property.
' Remarks: A Worksheet's CodeName property is distinct from its Name property.
'   The CodeName property is visible in the Project Explorer in the Excel
'   VBA IDE. The Name property is shown on the Worksheet tab in the regular
'   Excel window. The method searches the active workbook.
' Parameter(s):
'   pCodeName - Represents the value of the CodeName property of the sheet of
'     interest.
' Return(s): The worksheet object with the specifed CodeName property.
' Date Created: 2026-05-17
' Date Last Modified: 2026-06-09
'------------------------------------------------------------------------------'
Public Function GetWorksheetByCodeName(pCodeName As String) As Worksheet
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "GetWorksheetByCodeName"
  
  Dim sheetFound As Boolean
  Dim wb As Workbook
  Dim ws As Worksheet
  
  Set wb = ThisWorkbook
  
  sheetFound = False
  For Each ws In wb.Worksheets
      If UCase(ws.CodeName) = UCase(pCodeName) Then
        Set GetWorksheetByCodeName = ws
        sheetFound = True
        GoTo Exit_Proc
      End If
  Next ws
  
  If Not sheetFound Then
    Set GetWorksheetByCodeName = Nothing
  End If
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function