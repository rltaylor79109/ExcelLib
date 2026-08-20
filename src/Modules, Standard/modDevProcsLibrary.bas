Option Explicit
Public Const DEBUG_MODE As Boolean = True

'------------------------------------------------------------------------------'
' Module Name: modDevProcsLibrary
' Summary: Contains library (not application specific) methods used during
'   development of this workbook.
' Date Created: 2026-06-14
' Date Last Modified: 2026-08-20
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Fields
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Constant Fields
'------------------------------------------------------------------------------'

' The name of this module.
Private Const MODULE_NAME As String = "modDevProcsLibrary"

'------------------------------------------------------------------------------'
' Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Public Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Summary: Sets the horizontal aligment format of the current selection to
'   center across selection.
' Remarks: Did not work properly until the method cleared the center across
'   selection for before applying it.
' Date Created: 2026-05-14
' Date Last Modified: 2026-08-02
'------------------------------------------------------------------------------'
Public Sub CenterAcrossSelecton()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "CenterAcrossSelecton"
  
  Dim cell As Range
  Dim errorOccurred As Boolean
  Dim targetRng As Range

  errorOccurred = False
  Application.ScreenUpdating = False
  
  ' Check if user selected cells
  If TypeName(Selection) <> "Range" Then GoTo Exit_Proc
   
  Set targetRng = Selection
  
  ' Step 1: Clear xlCenterAcrossSelection ONLY within your current selection
  ' to break any lingering connections from previous formatting runs
  For Each cell In targetRng
      If cell.HorizontalAlignment = xlCenterAcrossSelection Then
        cell.HorizontalAlignment = xlGeneral
      End If
  Next cell
  
  ' Step 2: Apply Center Across Selection to the selected 4-column block
  targetRng.HorizontalAlignment = xlCenterAcrossSelection

Exit_Proc:
  Application.ScreenUpdating = True
  
  If errorOccurred Then
    ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  End If
  
  Exit Sub
Err_Proc:
  errorOccurred = True
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Determines if a Chartsheet has a code behind module.
' Parameter(s)
'   cs - The Chartsheet
' Return(s): True if the Chartsheet has a code behind module.
' Date Created: 2026-08-20
' Date Last Modified: 2026-08-20
'------------------------------------------------------------------------------'
Function ChartSheetHasCode(cs As Chart) As Boolean
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ChartSheetHasCode"

  Dim vbComp As Object
  Dim localCodeName As String
  
' Safely get localCodeName to prevent runtime failure
  On Error Resume Next
  localCodeName = cs.codeName
  On Error GoTo Err_Proc
  
  ' If localCodeName is empty, it has no accessible module or code
  If Len(Trim$(localCodeName)) = 0 Then GoTo Exit_Proc
  
  On Error Resume Next
  Set vbComp = cs.Parent.VBProject.VBComponents(localCodeName)
  On Error GoTo Err_Proc
  
  If Not vbComp Is Nothing Then
    ChartSheetHasCode = (vbComp.codeModule.CountOfLines > 0)
  End If

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Prints a list of the names of the chart sheets in this
'   workbook to debug output.
' Date Created: 2026-05-17
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Sub ListChartSheetNames()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ListChartSheetNames"

  Dim chtTitle As String
  Dim cs As Chart
  Dim csCount As Long
  Dim wb As Workbook
  
  Set wb = ThisWorkbook
  csCount = 0
  For Each cs In wb.Charts
      csCount = csCount + 1
      If cs.HasTitle Then
        chtTitle = cs.ChartTitle.text
      Else
        chtTitle = "[No Title]"
      End If
      Debug.Print cs.name & ", " & chtTitle
  Next cs
  
  If csCount = 0 Then
    Debug.Print "This workbook does not contain chart sheets."
  End If
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Prints a list of the CodeNames of Chartsheets with a non-empty code
'   behind module to debug output.
' Date Created: 2026-08-20
' Date Last Modified: 2026-08-20
'------------------------------------------------------------------------------'
Sub ListChartsheetsWithCode()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ListChartsheetsWithCode"

  Dim hasCode As Boolean
  Dim cs As Chart
  Dim wb As Workbook
  
  Set wb = ThisWorkbook
  For Each cs In wb.Charts
    hasCode = ChartSheetHasCode(cs)
    If hasCode Then
      Debug.Print cs.codeName & " has code"
    End If
  Next cs
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Prints a list of the names of all the embedded charts in this
'   workbook to debug output.
' Date Created: 2026-05-17
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Sub ListEmbeddedChartNames()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ListEmbeddedChartNames"

  Dim chtObj As ChartObject
  Dim chtTitle As String
  Dim chtCount As Long
  Dim output As String
  Dim wb As Workbook
  Dim ws As Worksheet
  
  Set wb = ThisWorkbook
  chtCount = 0
  For Each ws In wb.Worksheets
    For Each chtObj In ws.ChartObjects
      chtCount = chtCount + 1
      If chtObj.Chart.HasTitle Then
        chtTitle = chtObj.Chart.ChartTitle.text
      Else
        chtTitle = "[No Title]"
      End If
      
      output = ws.name & ", " & chtObj.name & ", " & chtTitle
      Debug.Print output
    Next chtObj
  Next ws
  
  If chtCount = 0 Then
    Debug.Print "This workbook does not contain any embedded charts."
  End If
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub
'------------------------------------------------------------------------------'
' Summary: Prints a list of the CodeNames of Worksheets with a non-empty code
'   behind module to debug output.
' Date Created: 2026-08-20
' Date Last Modified: 2026-08-20
'------------------------------------------------------------------------------'
Sub ListWorksheetsWithCode()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ListWorksheetsWithCode"

  Dim hasCode As Boolean
  Dim ws As Worksheet
  Dim wb As Workbook
  
  Set wb = ThisWorkbook
  For Each ws In wb.Worksheets
    hasCode = WorksheetHasCode(ws)
    If hasCode Then
      Debug.Print ws.codeName & " has code"
    End If
  Next ws
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Unhides all Worksheets and Chartsheets in this Workbook.
' Date Created: 2026-07-14
' Date Last Modified: 2026-07-14
'------------------------------------------------------------------------------'
Sub UnhideAllRowsInWorkbook()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "UnhideAllRowsInWorkbook"
  
  Const title As String = "Unhiding All Rows in Worksheets"
  
  Dim buttons As VbMsgBoxStyle
  Dim curActiveSheet As Object
  Dim errorOccurred As Boolean
  Dim prompt As String
  Dim wb As Workbook
  Dim ws As Worksheet
    
  errorOccurred = False
  
  OptimizeAppEnvForSpeed True
    
  Set wb = ThisWorkbook
  Set curActiveSheet = wb.ActiveSheet
  
  ' Loop through each worksheet in the active workbook
  For Each ws In wb.Worksheets
    ws.Rows.Hidden = False
  Next ws

Exit_Proc:
  OptimizeAppEnvForSpeed False
  
  If Not curActiveSheet Is Nothing Then
    curActiveSheet.Activate
  End If
  
  If errorOccurred Then
    ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
    buttons = vbExclamation
    prompt = "The was a problem unhidding all the rows."
  Else
    buttons = vbInformation
    prompt = "All Rows in all Worksheets have been unhidden."
  End If
  
  MsgBox prompt, buttons, title

  Exit Sub
Err_Proc:
  errorOccurred = True
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Unhides all Worksheets and Chartsheets in this Workbook.
' Date Created: 2017-03-24
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Sub UnhideAllSheetsAndCharts()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "UnhideAllSheetsAndCharts"
  
  Dim cht As Chart
  Dim wb As Workbook
  Dim ws As Worksheet
  
  Set wb = ThisWorkbook
  For Each ws In wb.Worksheets
    ws.Visible = xlSheetVisible
  Next ws
  
  For Each cht In wb.Charts
    cht.Visible = xlSheetVisible
  Next cht

Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Unhides worksheets based on user input.
' Date Created: 2017-03-24
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Sub UnhideSomeSheets()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "UnhideSomeSheets"
  
  Dim buttons As VbMsgBoxStyle
  Dim prompt As String
  Dim response As VbMsgBoxResult
  Dim wb As Workbook
  Dim ws As Worksheet
  Dim wsName As String
  
  Set wb = ThisWorkbook
  For Each ws In wb.Worksheets
    If ws.Visible = xlSheetHidden Then
      wsName = ws.name
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
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Determines if a Worksheet has a code behind module.
' Parameter(s)
'   ws - The Worksheet
' Return(s): True if the Worksheet has a code behind module.
' Date Created: 2026-08-20
' Date Last Modified: 2026-08-20
'------------------------------------------------------------------------------'
Function WorksheetHasCode(ws As Worksheet) As Boolean
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "WorksheetHasCode"

  Dim vbComp As Object
  
  On Error Resume Next
  Set vbComp = ws.Parent.VBProject.VBComponents(ws.codeName)
  On Error GoTo Err_Proc
  
  If Not vbComp Is Nothing Then
      WorksheetHasCode = (vbComp.codeModule.CountOfLines > 0)
  End If

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function
