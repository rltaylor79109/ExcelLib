'------------------------------------------------------------------------------'
' Summary: Gets the Name property from the Worksheet specified by CodeName.
' Remarks: A Worksheet's CodeName property is distinct from its Name property.
'   The CodeName property is visible in the Project Explorer in the Excel
'   VBA IDE. The Name property is shown on the Worksheet tab in the regular
'   Excel window. The method searches the active workbook.
' Parameter(s):
'   pCodeName - Represents the value of the CodeName property of the sheet of
'     interest.
' Return(s): The value of the Name property from the Worksheet specified by
  ' CodeName.
' Date Created: 2026-05-29
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Function GetWorksheetNameByCodeName(pCodeName As String) As String
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "GetWorksheetNameByCodeName"
  
  Dim ws As Worksheet
  
  Set ws = GetWorksheetByCodeName(pCodeName)
  If ws Is Nothing Then
    GetWorksheetNameByCodeName = "NOT FOUND"
  Else
    GetWorksheetNameByCodeName = ws.Name
  End If
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function