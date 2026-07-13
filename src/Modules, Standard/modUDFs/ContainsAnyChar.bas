'------------------------------------------------------------------------------'
' Summary: Determines if the specified string contains any characters
'   contained in another string.
' Parameter(s):
'   pText - The string searched to determine if it contains any of the
'     characters contained by the charsToFind string.
'   charsToFind - A string that represents the list of characters to find
'     in the pText parameter.
' Returns: True if the pText parameter string contains any characters
'   contained charsToFind parameter; otherwise, False.
' Date Created: 2026-05-18
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Function ContainsAnyChar(pText As String, charsToFind As String) As Boolean
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ContainsAnyChar"

  Dim i As Long
  Dim curChar As String
    
  ' Loop through each character in the "charsToFind" string
  For i = 1 To Len(charsToFind)
    curChar = Mid(charsToFind, i, 1)
        
  ' If the character is found anywhere in the text, return True and exit
  If InStr(1, pText, curChar, vbTextCompare) > 0 Then
    ContainsAnyChar = True
    GoTo Exit_Proc
    End If
  Next i

  ContainsAnyChar = False
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function