'------------------------------------------------------------------------------'
' Purpose:  Determines if the specified VBA component contains the specified
'   string.
' Parameter(s):
'   searchIn - The VBA component (Modules, Sheets, Forms, Classes) to search
'     in.
'   searchFor - The string to search for.
'   wholeWordOnly - If True, the search matches whole words only, so that
'     "Pie" matches "Apple Pie" but does not match "Pied Piper"; otherwise,
'     the match occurs whether the string is within a word or not
'     so that "Pie" matches both "Apple Pie" and "Pied Piper". It is an
'     optional parameter with a default value of False.
'   matchCase - If True, the search is case-sensitive; otherwise the
'     search is case-insensitive. It is an optional parameter with a default
'     value of False.
' Returns: True if the specified VBA component contains the specified
'   string; otherwise False.
' Date Created: 2026-06-16
' Date Last Modified: 2026-06-16
'------------------------------------------------------------------------------'
Public Function ComponentContainsString( _
  searchIn As VBComponent, _
  searchFor As String, _
  Optional wholeWordOnly As Boolean = False, _
  Optional matchCase As Boolean = False _
  ) As Boolean

  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ComponentContainsString"
  Dim codeMod As CodeModule
  Dim i As Long
  Dim lineText As String
  ComponentContainsString = False
  
  Set codeMod = searchIn.CodeModule
  For i = 1 To codeMod.CountOfLines
    lineText = codeMod.Lines(i, 1)
    If ContainsString( _
      searchIn:=lineText, _
      searchFor:=searchFor, _
      wholeWordOnly:=wholeWordOnly, _
      matchCase:=matchCase) Then

      ComponentContainsString = True
      Exit For
    End If
  Next i

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function