'------------------------------------------------------------------------------'
' Summary: Unhides worksheets based on user input.
' Date Created: 2017-03-24
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Sub UnhideSomeSheets()
  On Error GoTo Err_Proc
  Const methodName As String = "UnhideSomeSheets"
  
  Dim buttons As VbMsgBoxStyle
  Dim prompt As String
  Dim response As VbMsgBoxResult
  Dim wb As Workbook
  Dim ws As Worksheet
  Dim wsName As String
  
  Set wb = ThisWorkbook
  For Each ws In wb.Worksheets
    If ws.Visible = xlSheetHidden Then
      wsName = ws.Name
      prompt = "Unhide the following sheet?" _
        & vbNewLine & wsName
      buttons = vbYesNoCancel
      response = MsgBox(prompt, buttons)
      If response = vbYes Then
        ws.Visible = xlSheetVisible
      ElseIf response = vbCancel Then
        Exit For
      End If
    End If
  Next ws

Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, methodName
  Resume Ex