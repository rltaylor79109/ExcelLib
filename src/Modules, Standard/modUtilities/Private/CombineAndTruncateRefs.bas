'------------------------------------------------------------------------------'
' Summary: A helper function that parses a collection reference strings from
'   the Formulas Used for Data Valdiation Report table's References column.
' Remarks: The references strings have the format"
'   'sht 1'!$A$1,$A$2;'sht 2'!$B$2,$B$3,...
'   where the commas separate the individual cell addressess. If the string
'   ends in ",...", there are additional references not listed.
'   The example above would place these keys in the specified reference
'   dictionary, 'sht 1'!|$A$1, 'sht 1'!|$A$2, 'sht 2'!|$B$2, 'sht 2'!|$B$3
'   The the same cell reference is in more than one item in the collection,
'   it is not duplicated in the result.
' Parameter(s)
'   refsrefStrings - A collection reference strings from the Formulas Used
'     for Data Valdiation Report table's References column.
'   refCountLimit - Places a limit on the number of references that are added
'     to the resulting string. If the limit is exceeded, then "..." is
'     appended to the result. It is an optional parameter with a default
'     value of 25.
' Date Created: 2026-07-09
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Private Function CombineAndTruncateRefs( _
  ByVal refStrings As Collection, _
  Optional refCountLimit As Long = 25) _
  As String
  
  Const METHOD_NAME = "CombineAndTruncateRefs"
  
  Dim result As String
  Dim curWsName As String
  Dim item As Variant
  Dim key As Variant
  Dim parts() As String
  Dim refCount As Integer
  Dim refsDict As Dictionary
   
  Set refsDict = New Dictionary
  
  ' 1. Clean and parse every string in the collection into the refsDictionary
  For Each item In refStrings
    ParseReferences CStr(item), refsDict
  Next item
    
  ' 2. Rebuild the string from the unique reference dictionary keys, up to
  ' a maximum ofthe specified reference refCount limit.
  result = ""
  refCount = 0
  curWsName = ""
  For Each key In refsDict.Keys
    If refCount >= refCountLimit Then
      result = result & ",..."
      Exit For
    End If

    parts = Split(key, "|") ' parts(0) is Sheet, parts(1) is Cell
    
    ' If the worksheet changes, append the new worksheet prefix
    If parts(0) <> curWsName Then
      ' Append separator if this isn't the very first item in the result
      If result <> "" Then
        result = result & ";"
      End If
      curWsName = parts(0)
      result = result & curWsName & parts(1)
    Else
        ' Same sheet, just append the cell with a comma
        result = result & "," & parts(1)
    End If
    
    refCount = refCount + 1
  Next key
    
  ' Clean up any trailing semicolon if it ended exactly at the reference count
  ' limit without a ",..."
  If Right(result, 1) = ";" Then
    result = Left(result, Len(result) - 1)
  End If
  CombineAndTruncateRefs = result
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function
