Option Explicit
Option Private Module

'------------------------------------------------------------------------------'
' Module Name: modDevProcsLibrary
' Summary: Contains library (not application specific) methods used during
'   development of this workbook.
' Date Created: 2026-06-14
' Date Last Modified: 2026-09-07
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
Public Function ChartSheetHasCode(cs As Chart) As Boolean
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
    ChartSheetHasCode = (vbComp.CodeModule.CountOfLines > 0)
  End If

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Exports the lambda functions from the current workbook to test
'   files.
' Date Created: 2026-09-03
' Date Last Modified: 2026-09-07
'------------------------------------------------------------------------------'
Sub ExportLambdasToFiles()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ExportLambdasToFiles"
  
  Dim commentText As String
  Dim exportCount As Long
  Dim folderPath As String
  Dim folderPicker As FileDialog
  Dim formulaText As String
  Dim filePath As String
  Dim fileNum As Integer
  Dim lclName As Name
  Dim posComment As Long
  Dim isCommentInSrc As Boolean
  Dim isLambda As Boolean
  Dim posFcnName As Long
  
  ' 1. Prompt user to select output folder
  Set folderPicker = Application.FileDialog(msoFileDialogFolderPicker)
  With folderPicker
    .title = "Select Folder to Export Lambda Functions"
    .AllowMultiSelect = False
    If .Show = -1 Then
      folderPath = .SelectedItems(1)
    Else
      MsgBox "Export canceled.", vbInformation
      GoTo Exit_Proc
    End If
  End With
  
  If Right(folderPath, 1) <> "\" Then
    folderPath = folderPath & "\"
  End If
  exportCount = 0

  For Each lclName In ThisWorkbook.Names
    formulaText = lclName.RefersTo
    isLambda = InStr(1, formulaText, "LAMBDA(", vbTextCompare) > 0
    If isLambda Then
      filePath = folderPath & lclName.Name & ".xlsxfx"
      fileNum = FreeFile
      Open filePath For Output As #fileNum
      Print #fileNum, formulaText
      Close #fileNum
      exportCount = exportCount + 1
      End If ' isLambda
  Next lclName
  
  MsgBox _
    exportCount & " Lambda function(s) successfully exported to:" & vbCrLf & folderPath, _
    vbInformation, _
    "Lambda Function Export"

Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Exports all VBA code components (Standard, Class, Worksheet,
'   Chartsheet, and UserForm) from the current workbook to text files.
' Remarks: Requires access to the VBA Object Model.
'   (Excel Options > Trust Center > Trust Center Settings > Macro Settings >
'   Check "Trust access to the VBA project object model")
' Date Created: 2026-09-07
' Date Last Modified: 2026-09-07
'------------------------------------------------------------------------------'
Public Sub ExportVBAModulesToFiles()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ExportVBAModulesToFiles"
  
  Dim exportCount As Long
  Dim ext As String
  Dim folderPath As String
  Dim folderPicker As FileDialog
  Dim filePath As String
  Dim vbComp As Object ' VBComponent
  
  ' 1. Prompt user to select output folder
  Set folderPicker = Application.FileDialog(msoFileDialogFolderPicker)
  With folderPicker
    .title = "Select Folder to Export VBA Modules"
    .AllowMultiSelect = False
    If .Show = -1 Then
      folderPath = .SelectedItems(1)
    Else
      MsgBox "Export canceled.", vbInformation
      GoTo Exit_Proc
    End If
  End With ' folderPicker
  
  If Right(folderPath, 1) <> "\" Then
    folderPath = folderPath & "\"
  End If
  exportCount = 0
  
  ' 2. Loop through all VBA components in the workbook
  For Each vbComp In ThisWorkbook.VBProject.VBComponents
    ' Determine file extension based on module component type
    Select Case vbComp.Type
        Case 1 ' vbext_ct_StdModule (.bas)
        ext = ".bas"
      Case 2 ' vbext_ct_ClassModule (.cls)
        ext = ".cls"
      Case 3 ' vbext_ct_MSForm (.frm)
        ext = ".frm"
      Case 100 ' vbext_ct_Document (Worksheets, ChartSheets, ThisWorkbook) (.cls)
        ext = ".cls"
      Case Else
        ext = ".txt"
    End Select

    ' Only export components that contain code lines
    If vbComp.CodeModule.CountOfLines > 0 Then
      filePath = folderPath & vbComp.Name & ext
      vbComp.Export filePath
      exportCount = exportCount + 1
    End If
  
  Next vbComp
  
  MsgBox _
  exportCount & " VBA module(s) successfully exported to:" & vbCrLf & folderPath, _
  vbInformation, _
  "VBA Module Export"
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
  End Sub

'------------------------------------------------------------------------------'
' Summary: Prints a list of all cells that spill their data.
' Date Created: 2026-08-29
' Date Last Modified: 2026-08-29
'------------------------------------------------------------------------------'
Public Sub FindAllSpillCells()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "FindAllSpillCells"
  
  Dim cell As Range
  Dim wb As Workbook
  Dim ws As Worksheet
  
  Set wb = ActiveWorkbook
  For Each ws In wb.Worksheets
    On Error Resume Next
    For Each cell In ws.UsedRange.SpecialCells(xlCellTypeFormulas)
      If cell.HasSpill Then
        Debug.Print "'" & ws.Name & "'!" & cell.Address
      End If
    Next cell
    On Error GoTo Err_Proc
  Next ws
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Imports Lambda functions from files.
' Remarks: NEEDS TESTING
' Date Created: 2026-09-04
' Date Last Modified: 2026-09-06
'------------------------------------------------------------------------------'
Public Sub ImportLambdasFromFiles()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ImportLambdasFromFiles"

  Const msgBoxTitle As String = "Lambda Function Import"
  
  Dim errorOccurred As Boolean
  Dim defName As Name
  Dim fcnDef As String
  Dim fcnComment As String
  Dim fcnName As String
  Dim fDialog As FileDialog
  Dim fileItem As Variant
  Dim fileSysObj As FileSystemObject
  Dim txtStrm As TextStream

  ' Get the files from the user.
  Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
  With fDialog
    .title = "Select Lambda Function Files to Import"
    .Filters.Clear
    .Filters.Add "Text Files", "*.txt; *.lambda; *.xlsxfx"
    .Filters.Add "All Files", "*.*"
    .AllowMultiSelect = True
    If .Show <> -1 Then
      Debug.Print
      MsgBox _
        "No files were selected for importing", _
        vbOK + vbInformation, _
        msgBoxTitle
      GoTo Exit_Proc
    End If
  End With
    
  OptimizeAppEnvForSpeed True
    
  ' Initialize FileSystemObject for reading text files
  Set fileSysObj = New FileSystemObject
    
  For Each fileItem In fDialog.SelectedItems
    ' Read file text into string
    Set txtStrm = fileSysObj.OpenTextFile(CStr(fileItem), 1) ' 1 = ForReading
    fcnDef = txtStrm.ReadAll
    txtStrm.Close
    
    fcnName = ExtractLambdaFcnName(fcnDef) ' get the embedded function name.
    If fcnName = "" Then
      Debug.Print "Failed to import function from """ & CStr(fileItem) & _
        """- The function definition does not contain an embedded function name."
      GoTo Next_FileItem
    End If
    
    fcnComment = ExtractLambdaComment(fcnDef) ' get the embedded function comment.
    
    ' Add or overwrite the defined name in the active workbook
    On Error Resume Next
    Set defName = ThisWorkbook.Names.Add(Name:=fcnName, RefersTo:=fcnDef)
    If err.Number <> 0 Then
      Debug.Print "Failed to import: " & fcnName & " - " & err.Description
      err.Clear
    Else
      defName.Comment = fcnComment
      Debug.Print "Successfully imported: " & fcnName & " from """ & CStr(fileItem) & """."
    End If
    On Error GoTo Err_Proc
    
Next_FileItem:
  Next fileItem

Exit_Proc:
  OptimizeAppEnvForSpeed False
  If errorOccurred Then
    ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
    MsgBox _
      "There was a problem importing the Lambda function(s).", _
      vbOK + vbCritical, _
      msgBoxTitle
  Else
    MsgBox _
      "All the Lambda function(s) were successfully imported.", _
      vbOK + vbInformation, _
      msgBoxTitle
  End If
  Exit Sub
Err_Proc:
  errorOccurred = True
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Prints a list of the names of the chart sheets in this
'   workbook to debug output.
' Date Created: 2026-05-17
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Sub ListChartSheetNames()
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
      Debug.Print cs.Name & ", " & chtTitle
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
Public Sub ListChartsheetsWithCode()
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
Public Sub ListEmbeddedChartNames()
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
      
      output = ws.Name & ", " & chtObj.Name & ", " & chtTitle
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
Public Sub ListWorksheetsWithCode()
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
Public Sub UnhideAllRowsInWorkbook()
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
Function WorksheetHasCode(ByVal ws As Worksheet) As Boolean
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "WorksheetHasCode"

  Dim vbComp As Object
  
  On Error Resume Next
  Set vbComp = ws.Parent.VBProject.VBComponents(ws.codeName)
  On Error GoTo Err_Proc
  
  If Not vbComp Is Nothing Then
      WorksheetHasCode = (vbComp.CodeModule.CountOfLines > 0)
  End If

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Private Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Summary: Extracts the comment embedded in a Lambda function definition
'   formatted as _comment, N("The comment.") from the specified string.
' Parameter(s):
'   lambdaFcnDefStr - The string representing the Lambda function definition.
' Return(s): If found, the comment embedded in the specified string that
'   represents a Lambda function definition containing a comment formatted as
'   _comment, N("the comment"); otherwise, if no comment is found, "",
'   the empty string.
' Date Created: 2026-09-05
' Date Last Modified: 2026-09-05
'------------------------------------------------------------------------------'
Private Function ExtractLambdaComment(ByVal lambdaFcnDefStr As String) As String
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ExtractLambdaComment"
  
  ' Matches _comment, N("The comment.") across spaces and newlines
  Const targetPattern As String = "_comment\s*,\s*N\s*\(\s*""([^""]+)""\s*\)"
  
  Dim matchColl As MatchCollection
  Dim regEx As RegExp

  Set regEx = New RegExp
  With regEx
    .Global = False
    .IgnoreCase = True
    .Multiline = True
    .Pattern = targetPattern
  End With
  
  If regEx.Test(lambdaFcnDefStr) Then
    Set matchColl = regEx.Execute(lambdaFcnDefStr)
    ' Return the first capture group (the string inside the quotes)
    ExtractLambdaComment = matchColl(0).SubMatches(0)
  Else
    ExtractLambdaComment = ""
  End If
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Extracts the function name embedded in a Lambda function definition
'   formatted as _functionName, N("The Function Name") from the specified
'   string.
' Parameter(s):
'   lambdaFcnDefStr - The string representing the Lambda function definition.
' Return(s): If found, the function name embedded in the specified string that
'   represents a Lambda function definition containing a comment formatted as
'   _functionName, N("The Function Name"); otherwise, if no comment is found,
'   "", the empty string.
' Date Created: 2026-09-05
' Date Last Modified: 2026-09-05
'------------------------------------------------------------------------------'
Private Function ExtractLambdaFcnName(ByVal lambdaFcnDefStr As String) As String
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ExtractLambdaFcnName"
  
  ' Matches _functionName, N("captured_text") across spaces and newlines
  Const targetPattern As String = _
    "_functionName\s*,\s*N\s*\(\s*""([^""]+)""\s*\)"
  
  Dim matchColl As MatchCollection
  Dim regEx As RegExp

  Set regEx = New RegExp
  With regEx
    .Global = False
    .IgnoreCase = True
    .Multiline = True
    .Pattern = targetPattern
  End With
  
  If regEx.Test(lambdaFcnDefStr) Then
    Set matchColl = regEx.Execute(lambdaFcnDefStr)
    ' Return the first capture group (the string inside the quotes)
    ExtractLambdaFcnName = matchColl(0).SubMatches(0)
  Else
    ExtractLambdaFcnName = ""
  End If
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function
