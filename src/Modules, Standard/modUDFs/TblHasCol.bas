'------------------------------------------------------------------------------'
' Summary: Determines if the specified table the contains column (field)
'   specified by column heading.
' Parameter(s):
'   pTblLstObj - The table to search.
'   pColHeading - The column heading to search for.
' Returns - True if the specified table contains the column (field) specified
'   by column heading; Otherwise, False
' Date Created: 2026-08-15
' Date Last Modified: 22026-08-15
'------------------------------------------------------------------------------'
Public Function TblHasCol( _
  ByVal pTblLstObj As ListObject, _
  ByVal pColHeading As String _
  ) As Boolean
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "TblHasCol"
  
  Dim listCol As ListColumn
  
  If pTblLstObj Is Nothing Or Trim(pColHeading) = "" Then
    err.Raise _
      Number:=VBA_ERR_INVALID_PROCEDURE_CALL_OR_ARGUMENT, _
      Source:=MODULE_NAME & "." & METHOD_NAME, _
      Description:="Invalid table list object or empty column heading"
      GoTo Exit_Proc
  End If
  
  On Error Resume Next
  Set listCol = pTblLstObj.ListColumns(pColHeading)
  On Error GoTo Err_Proc
    
  TblHasCol = Not listCol Is Nothing
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function