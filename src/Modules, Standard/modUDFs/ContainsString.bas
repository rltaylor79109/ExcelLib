'------------------------------------------------------------------------------'
' Purpose:  Determines if one string contains another string search for whole
'   words only.
' Parameter(s):
'   searchIn - The main string you want to search inside of.
'   searchFor - The string you are looking for.
'   wholeWordOnly - If True, the searching looks for whole words only. It is
'     an optional parameter with a default value of False.
'   matchCase - If true the search is case-sensitive; otherwise the
'     search is case-insensitive. It is an optional parameter with a default
'     value of False.
' Returns: True if searchIn contains the whole wordsearchFor; otherwise False.
' Date Created: 2026-06-15
' Date Last Modified: 2026-06-15
'------------------------------------------------------------------------------'
Public Function ContainsString( _
  ByVal searchIn As String, _
  ByVal searchFor As String, _
  Optional ByVal wholeWordOnly As Boolean = False, _
  Optional ByVal matchCase As Boolean = False _
  ) As Boolean
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ContainsString"
  
  Dim escapedSearchFor As String
  Dim finalPattern As String
  Dim regEx As RegExp

  Set regEx = New RegExp
  
  ' Quick exit if either string is empty
  If Len(searchIn) = 0 Or Len(searchFor) = 0 Then
      ContainsString = False
      GoTo Exit_Proc
  End If
  
  escapedSearchFor = EscapeRegexChars(searchFor)
  
  If wholeWordOnly Then
      finalPattern = "\b" & escapedSearchFor & "\b"
  Else
      finalPattern = escapedSearchFor
  End If
  
  With regEx
      .Pattern = finalPattern
      .IgnoreCase = Not matchCase
      .Global = False
  End With
  
  ContainsString = regEx.test(searchIn)
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function