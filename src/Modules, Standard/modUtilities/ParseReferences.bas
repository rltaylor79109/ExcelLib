'------------------------------------------------------------------------------'
' Summary: A helper function that parses references string the Formulas
'   Used for Data Valdiation Report table's References column.
' Remarks: The references strings have the format"
'   'sht 1'!$A$1,$A$2;'sht 2'!$B$2,$B$3,...
'   where the commas separate the individual cell addressess. If the string
'   ends in ",...", there are additional references not listed.
'   The example above would place these keys in the specified reference
'   dictionary, 'sht 1'!|$A$1, 'sht 1'!|$A$2, 'sht 2'!|$B$2, 'sht 2'!|$B$3
' Parameter(s)
'   refsStr - The reference string to be parsed.
'   refsDict - The dictionary where the parsed data is placed. It is an output
'     parameter.
' Date Created: 2026-07-09
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Private Sub ParseReferences( _
  ByVal refsStr As String, _
  ByRef refsDict As Dictionary)
  
  Const METHOD_NAME = "ParseReferences"
    
  Dim block As String
  Dim cellAddr As String
  Dim cellIdx As Integer
  Dim cellAddrList() As String
  Dim cellsPart As String
  Dim exclamPos As Integer
  Dim strBuilder As String
  Dim wsBlockIdx As Integer
  Dim wsBlocks() As String
  Dim wsName As String
  Dim uniqueKey As String
  
  ' Remove trailing truncation markers
  strBuilder = Replace(refsStr, "...", "")
  If Trim(strBuilder) = "" Then
    refsStr = ""
    GoTo Exit_Proc
  End If
  
  ' Split into worksheet wsBlocks by semicolon
  wsBlocks = Split(strBuilder, ";")
  
  For wsBlockIdx = LBound(wsBlocks) To UBound(wsBlocks)
    block = Trim(wsBlocks(wsBlockIdx))
    
    If block = "" Then
      GoTo Next_wsBlockIdx_Iter
    End If
    
    ' Extract the Worksheet name part (everything up to and including the "!")
    exclamPos = InStr(block, "!")
    If exclamPos > 0 Then
      wsName = Left(block, exclamPos)
      cellsPart = Mid(block, exclamPos + 1)
    Else
        wsName = ""
        cellsPart = block
    End If
    
    ' Split the cell addresses by comma
    cellAddrList = Split(cellsPart, ",")
    
    For cellIdx = LBound(cellAddrList) To UBound(cellAddrList)
      cellAddr = Trim(cellAddrList(cellIdx))
      If cellAddr <> "" Then
        ' Create a unique tracking key combining sheet and cell
        uniqueKey = wsName & "|" & cellAddr
        ' refsDictionary automatically ignores duplicates this way
        If Not refsDict.Exists(uniqueKey) Then
          refsDict.Add uniqueKey, True
        End If
      End If
    Next cellIdx
Next_wsBlockIdx_Iter:
  Next wsBlockIdx

Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub
