'------------------------------------------------------------------------------'
' Summary: Protects all of the sheets and charts in the Workbook.
' Remarks: Requires reference to "Microsoft Scripting Runtime."
' Parameter(s):
'   excludewsNamesDict - A dictionary that contains the names of the worksheets
'     to not protect. It is an optional parameter with a default value of
'     nothing.
'   excludeChrtNamesDict - A dictionary that contains the names of the worksheets
'     to not protect.
'   silentMode - If True, the method does not write messages to Debug.Print
'     or display a message to the user when complete. It is an optional
'     parameter with a default value of False.
' Date Created: 2026-02-04
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Public Sub ProtectAllSheetsAndCharts( _
  Optional ByVal excludewsNamesDict As Scripting.Dictionary = Nothing, _
  Optional ByVal excludeChrtNamesDict As Scripting.Dictionary = Nothing, _
  Optional ByVal silentMode = False)
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ProtectAllSheetsAndCharts"
  
  Dim cht As Chart
  Dim curActiveSheet As Object
  Dim excludedCht As Boolean
  Dim excludedWs As Boolean
  Dim wb As Workbook
  Dim ws As Worksheet
  
  Set wb = ThisWorkbook
  Set curActiveSheet = wb.ActiveSheet
    
  For Each ws In wb.Worksheets
    excludedWs = False
    If Not (excludewsNamesDict Is Nothing) Then
      excludedWs = excludewsNamesDict.Exists(ws.Name)
    End If
    
    If excludedWs Then
      If Not silentMode Then
        Debug.Print "Excluding " & ws.Name & " from Protection"
      End If
    Else
      If Not silentMode Then
        Debug.Print "Protecting " & ws.Name
      End If
      ws.Protect
    End If
  Next
  
  For Each cht In wb.Charts
    excludedCht = False
    If Not (excludeChrtNamesDict Is Nothing) Then
      excludedCht = excludeChrtNamesDict.Exists(cht.Name)
    End If
    
    If excludedCht Then
      If Not silentMode Then
        Debug.Print "Excluding " & cht.Name & " from Protection"
      End If
    Else
      If Not silentMode Then
        Debug.Print "Protecting " & cht.Name
      End If
      cht.Protect
    End If
  Next

Exit_Proc:
  If Not curActiveSheet Is Nothing Then
    curActiveSheet.Activate
  End If

  If Not silentMode Then
    Const prompt As String = "Protection complete."
    Const buttons As Long = vbInformation ' compiler fails if defined as vbMsgBoxStyle
    Const title As String = "Protecting Sheets and Charts"
    MsgBox prompt, buttons, title
  End If
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub