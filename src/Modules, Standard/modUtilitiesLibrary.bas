Option Explicit

'------------------------------------------------------------------------------'
' Module Name: modUtilitiesLibrary
' Summary: Contains library (not application specific)utility methods.
' Date Created: 2026-05-07
' Date Last Modified: 2026-08-18
'------------------------------------------------------------------------------'
'------------------------------------------------------------------------------'
' Fields
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Constant Fields
'------------------------------------------------------------------------------'

' The name of this module.
Private Const MODULE_NAME As String = "modUtilitiesLibrary"

'------------------------------------------------------------------------------'
' Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Public Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Summary: Get the standard prompt for a method error message box.
' Parameter(s):
'   err - The object that contains the error information
' Return(s): The standard prompt for a method error message box.
' Date Created: 2026-04-19
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Function GetErrorMsgBoxPrompt(err As ErrObject)
  GetErrorMsgBoxPrompt = _
    "Error: " & err.Number & vbCrLf & err.Description
End Function

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
      excludedWs = excludewsNamesDict.Exists(ws.name)
    End If
    
    If excludedWs Then
      If Not silentMode Then
        Debug.Print "Excluding " & ws.name & " from Protection"
      End If
    Else
      If Not silentMode Then
        Debug.Print "Protecting " & ws.name
      End If
      ws.Protect
    End If
  Next
  
  For Each cht In wb.Charts
    excludedCht = False
    If Not (excludeChrtNamesDict Is Nothing) Then
      excludedCht = excludeChrtNamesDict.Exists(cht.name)
    End If
    
    If excludedCht Then
      If Not silentMode Then
        Debug.Print "Excluding " & cht.name & " from Protection"
      End If
    Else
      If Not silentMode Then
        Debug.Print "Protecting " & cht.name
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

'------------------------------------------------------------------------------'
' Summary: Reads the contents of a text file and returns it as a string.
' Parameter(s):
'   filePath - The full path to the text file to read.
' Return(s): The full text contents of the specified file as a String.
' Date Created: 2026-07-21
' Date Last Modified: 2026-07-21
'------------------------------------------------------------------------------'
Public Function ReadTextFile(filePath As String) As String
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ReadTextFile"
  
  Dim fileNum As Integer
  Dim fileContent As String
  
  fileNum = FreeFile
  Open filePath For Input As #fileNum
  fileContent = Input$(LOF(fileNum), fileNum)
  Close #fileNum
  
  ReadTextFile = fileContent

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Runs the ProtectAllSheetsAndCharts using the default parameters.
' Remarks: It is a wrapper that allows this method to be assigned as the macro
'   for a button on a worksheet
' Date Created: 2026-05-08
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Sub RunProtectAllSheetsAndCharts()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunProtectAllSheetsAndCharts"

  ProtectAllSheetsAndCharts

Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Runs the UpsertDefNamesAllRefsRpt with the fastMode parameter set to
'   True so that the defined names worksheet references report is not updated
'   before the defined names all reference report is updated. This is
'   dramatically faster.
' Remarks: It is a wrapper that allows this method to be assigned as the macro
'   for a button on a worksheet.
' Date Created: 2026-06-17
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Sub RunUpsertDefNamesAllRptFast()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunUpsertDefNamesAllRptFast"

  UpsertDefNamesAllRefsRpt fastMode:=True
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Runs the UpsertDefNamesAllRefsRpt with the fastMode parameter set to
'   False so that the defined names worksheet references report is updated
'   before the defined names all reference report is updated. This is
'   dramatically slower.
' Remarks: It is a wrapper that allows this method to be assigned as the macro
'   for a button on a worksheet.
' Date Created: 2026-06-17
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Sub RunUpsertDefNamesAllRptSlow()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunUpsertDefNamesAllRptSlow"

  UpsertDefNamesAllRefsRpt fastMode:=False
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Runs UpsertDvAllRpt with the silent parameter
'   set to False. An update complete message is displayed to the user.
' Remarks: It is a wrapper that allows this method to be assigned as the macro
'   for a button on a worksheet.
' Date Created: 2026-07-07
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Public Sub RunUpsertDvAllRptVrbse()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunUpsertDvAllRptVrbse"

  UpsertDvAllRpt silent:=False
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Runs UpsertDvByFormulaRpt with the silent parameter
'   set to False. An update complete message is displayed to the user.
' Remarks: It is a wrapper that allows this method to be assigned as the macro
'   for a button on a worksheet.
' Date Created: 2026-07-07
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Public Sub RunUpsertDvByFormulaVrbse()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunUpsertDvByFormula"

  UpsertDvByFormulaRpt silent:=False
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Runs the UpsertDataValAuditRptVerbose with the silent parameter
'   set to True. Other than error messages, no user messages are displayed
'   upon completion of the update.
' Remarks: It is a wrapper that allows this method to be assigned as the macro
'   for a button on a worksheet.
' Date Created: 2026-06-17
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Public Sub RunUpsrtDataValRptSilent()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunUpsrtDataValRptSilent"

  UpsertDataValAuditRpt silent:=True
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Runs the UpsertDataValAuditRptVerbose with the silent parameter
'   set to False. An update complete message is displayed to the user.
' Remarks: It is a wrapper that allows this method to be assigned as the macro
'   for a button on a worksheet.
' Date Created: 2026-06-17
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Public Sub RunUpsrtDataValRptVrbse()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunUpsrtDataValRptVrbse"

  UpsertDataValAuditRpt silent:=False
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Runs the UpsertDefNamesWsRefRpt with the fastMode parameter set to
'   true. This is dramatically faster but the defined names references and
'   references counts are not shown in the report. The silent parameter is
'   set to False so that the usual update complete message is displayed to
'   the user.
' Remarks: It is a wrapper that allows this method to be assigned as the macro
'   for a button on a worksheet.
' Date Created: 2026-05-13
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Public Sub RunUpsrtDfNmsWsRptFastVrbs()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunUpsrtDfNmsWsRptFastVrbs"

  UpsertDefNamesWsRefRpt fastMode:=True, silent:=False
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Runs the UpsertDefNamesWsRefRpt with the fastMode parameter set to
'   false so that the defined names references and references count are shown
'   in the report. This is dramatically slower.  The silent parameter is
'   set to False so that the usual update complete message is displayed to
'   the user.
' Remarks: It is a wrapper that allows this method to be assigned as the macro
'   for a button on a worksheet.
' Date Created: 2026-07-17
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Public Sub RunUpsrtDfNmsWsRptSlowVrbs()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunUpsrtDfNmsWsRptSlowVrbs"

  UpsertDefNamesWsRefRpt fastMode:=False, silent:=False
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Runs the UpsertTblsWsRefRpt with the fastMode parameter set to
'   true. This is dramatically faster but the table references and
'   references counts are not shown in the report. The silent parameter is
'   set to False so that the usual update complete message is displayed to
'   the user.
' Remarks: It is a wrapper that allows this method to be assigned as the macro
'   for a button on a worksheet.
' Date Created: 2026-07-17
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Public Sub RunUpsrtTblsWsRptFastVrbs()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunUpsrtTblsWsRptFastVrbs"

  UpsertTblsWsRefRpt fastMode:=True, silent:=False
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Runs the UpsertDefNamesWsRefRpt with the fastMode parameter set to
'   false so that the table references and references count are shown
'   in the report. This is dramatically slower.  The silent parameter is
'   set to False so that the usual update complete message is displayed to
'   the user.
' Remarks: It is a wrapper that allows this method to be assigned as the macro
'   for a button on a worksheet.
' Date Created: 2026-07-17
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Public Sub RunUpsrtTblsWsRptSlowVrbs()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunUpsrtTblsWsRptSlowVrbs"

  UpsertTblsWsRefRpt fastMode:=False, silent:=False
  
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Optimize the Excel application environment for execution of heavy
'   procedure execution.
' Parameter(s):
'   optimize - If True, the Excel application environment is optimized
'     execution of heavy procedure execution; otherwise, the environment is
'     restored to Excel application defaults.
' Date Created: 2026-08-18
' Date Last Modified: 2026-08-18
'------------------------------------------------------------------------------'
Public Sub OptimizeAppEnvForSpeed(ByVal optimize As Boolean)
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "UnprotectAllSheetsAndCharts"
  
  With Application
    .ScreenUpdating = Not optimize
    .DisplayAlerts = Not optimize
    .EnableEvents = Not optimize
    .Calculation = IIf(optimize, xlCalculationManual, xlCalculationAutomatic)
    Application.Calculate
  End With

Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub


'------------------------------------------------------------------------------'
' Summary: Shows the standard method error message box.
' Parameter(s):
'   err - The object that contains the error information.
'   moduleName - The name of the module that contains the method that produced
'     the error.
'   METHOD_NAME - The name of the method that produced the error.
' Date Created: 2026-04-19
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Public Sub ShowMethodErrorMsgBox( _
  err As ErrObject, _
  moduleName As String, _
  METHOD_NAME As String)

  Dim prompt As String
  Dim buttons As Long
  Dim title As String
  
  prompt = "Source: Method " & moduleName & "." & METHOD_NAME & vbCrLf & _
    "Error " & err.Number & ": " & err.Description
  buttons = VbMsgBoxStyle.vbExclamation
  title = moduleName & "." & METHOD_NAME
  MsgBox prompt, buttons, title
End Sub

'------------------------------------------------------------------------------'
' Summary: Shows the standard property error message box.
' Parameter(s):
'   err - The object that contains the error information.
'   moduleName - The name of the module that contains the property that produced
'     the error.
'   propertyName - The name of the property that produced the error.
'   callType - The property call type.
' Date Created: 2026-04-19
' Date Last Modified: 2026-07-18
'------------------------------------------------------------------------------'
Public Sub ShowPropertyErrorMsgBox( _
  err As ErrObject, _
  moduleName As String, _
  propertyName As String, _
  callType As VbCallType)
  
  Dim buttons As VbMsgBoxStyle
  Dim callTypeString As String
  Dim prompt As String
  Dim title As String
  
  callTypeString = VbCallTypeToString(callType)
  prompt = "Source: Property " & moduleName & "." & propertyName & "." & _
     callTypeString & vbCrLf & _
    "Error " & err.Number & ": " & err.Description
  buttons = VbMsgBoxStyle.vbExclamation
  title = moduleName & "." & propertyName & "[Property" & callTypeString & "]"
  MsgBox prompt, buttons, title
End Sub

'------------------------------------------------------------------------------'
' Summary: Unprotects all of the sheets and charts in the Workbook.
' Parameter(s)
'   silentMode - If True, the method does not write messages to Debug.Print
'     or display a message to the user when complete. It is an optional
'     parameter with a default value of False.
' Date Created: 2026-03-12
' Date Last Modified: 2026-08-18
'------------------------------------------------------------------------------'
Public Sub UnprotectAllSheetsAndCharts(Optional silentMode As Boolean = False)
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "UnprotectAllSheetsAndCharts"

  Const title As String = "Unprotecting All Sheets and Charts"
  
  Dim buttons As VbMsgBoxStyle
  Dim cht As Chart
  Dim curActiveSheet As Object
  Dim errorOccurred As Boolean
  Dim prompt As String
  Dim wb As Workbook
  Dim ws As Worksheet
  
  errorOccurred = False
  
  ' Optimization: Turn off UI updates
  OptimizeAppEnvForSpeed True

  Set wb = ThisWorkbook
  Set curActiveSheet = wb.ActiveSheet
  
  For Each ws In wb.Worksheets
    If Not silentMode Then
      Debug.Print "Unprotecting " & ws.name
    End If
    ws.Unprotect
  Next
  
  For Each cht In wb.Charts
    If Not silentMode Then
      Debug.Print "Unprotecting " & cht.name
    End If
    cht.Unprotect
  Next
  
Exit_Proc:
  OptimizeAppEnvForSpeed False
  
  If Not curActiveSheet Is Nothing Then
    curActiveSheet.Activate
  End If
  
  If errorOccurred Then
    ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
    buttons = vbExclamation
    prompt = "The was a problem with the unprotection."
  Else
    buttons = vbExclamation
    prompt = "Unprotection complete."
  End If

  If Not silentMode Then
    MsgBox prompt, buttons, title
  End If

  Exit Sub
Err_Proc:
  errorOccurred = True
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Creates or updates a worksheet that show all the data validation
'   sources, their types, and the cell that used them, formatted as an Excel
'   table.
' Parameter(s)
'   silent - It True, the method does not display any messages to the user
'     except for errors; otherwise a report updated message is displayed to
'     the user when complete. It is an optional parameter with a default value
'     of False.
' Date Created: 2026-04-06
' Date Last Modified: 2026-08-18
'------------------------------------------------------------------------------'
Public Sub UpsertDataValAuditRpt(Optional silent As Boolean = False)
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "UpsertDataValAuditRpt"

  Const BUTTON_NAME As String = "btnUpsertDataValAuditRpt"
  Const COL_INDEX_FORMULA_SOURCE As Long = 1
  Const COL_INDEX_SHEET_INDEX As Long = 2
  Const COL_INDEX_CELL_ADDRESS As Long = 3
  Const COL_INDEX_VALIDATION_TYPE As Long = 4
  Const COL_NAME_FORMULA_SOURCE As String = "Formula_Source"
  Const COL_NAME_SHEET_NAME As String = "Sheet_Name"
  Const COL_NAME_CELL_ADDRESS As String = "Cell_Address"
  Const COL_NAME_VALIDATION_TYPE As String = "Validation_Type"
  Const COL_COUNT As Long = 4
  Const REPORT_WORKSHEET_CODENAME As String = _
    "SheetDataValidationAuditRpt"
  Const REPORT_WORKSHEET_NAME = "Data Validation Audit Rpt"
  Const REPORT_WORKSHEET_TITLE = "Data Validation Audit Report"
  Const TABLE_FIRST_ROW As Long = 3
  Const TABLE_NAME As String = "tbl_Data_Validation_Audit_Rpt"
  
  ' Valid XlDVType enumeration values ranges from 0 to 7
  Const XLDVTYPE_UNDEFINED As Long = -1
  
  Dim creatingNewWs As Boolean
  Dim errorOccurred As Boolean
  Dim ps As clsProtectionState
  Dim rowCount As Long
  Dim rptTable As ListObject
  Dim rptWs As Worksheet
  Dim targetCell As Range
  Dim tblRange As Range
  Dim updateButton As Button
  Dim validationRange As Range
  Dim vType As XlDVType
  Dim ws As Worksheet
  Dim wb As Workbook
  
  errorOccurred = False

  OptimizeAppEnvForSpeed True
  
  Set wb = ThisWorkbook
  
  ' Create or clear the audit sheet
  Set rptWs = _
    GetWsByCodeName(REPORT_WORKSHEET_CODENAME)
  creatingNewWs = (rptWs Is Nothing)
  
  If creatingNewWs Then
    Set rptWs = wb.Worksheets.Add(After:=Sheets(Sheets.count - 1))
    ThisWorkbook.VBProject.VBComponents(rptWs.codeName).name = _
      REPORT_WORKSHEET_CODENAME
    
    With rptWs
      .name = REPORT_WORKSHEET_NAME
      
      With .Cells.Font
        .name = "Aptos Narrow"
        .Size = 10
      End With
      
      With .Cells(1, 1)
        .value = REPORT_WORKSHEET_TITLE
        With .Font
          .name = "Aptos Display"
          .Size = 12
          .Bold = True
        End With
      End With
      
    End With
          
    ' Create Button: Left, Top, Width, Height (Positioned near Column F)
    Set updateButton = rptWs.buttons.Add( _
      Left:=300, _
      Top:=5, _
      Width:=80, _
      Height:=20)
    DoEvents ' Brief pause to let Excel register the object
    With updateButton
      .OnAction = "RunUpsrtDataValRptVrbse"
      .Caption = "Update"
      .name = BUTTON_NAME
      .Placement = xlFreeFloating
    End With
  Else
    Set ps = clsProtectionStateStatic.Create(rptWs)
    rptWs.Unprotect
    ' Clear existing ListObject if it exists to avoid conflicts
    On Error Resume Next
    rptWs.ListObjects(TABLE_NAME).Delete
    err.Clear
    On Error GoTo Err_Proc
  End If
    
  ' Headers\
  rptWs.Cells(TABLE_FIRST_ROW, COL_INDEX_FORMULA_SOURCE).value = _
    COL_NAME_FORMULA_SOURCE
  rptWs.Cells(TABLE_FIRST_ROW, COL_INDEX_SHEET_INDEX).value = _
    COL_NAME_SHEET_NAME
  rptWs.Cells(TABLE_FIRST_ROW, COL_INDEX_CELL_ADDRESS).value = _
    COL_NAME_CELL_ADDRESS
  rptWs.Cells(TABLE_FIRST_ROW, COL_INDEX_VALIDATION_TYPE).value = _
    COL_NAME_VALIDATION_TYPE
  rptWs.Range( _
    rptWs.Cells(TABLE_FIRST_ROW, 1), _
    rptWs.Cells(TABLE_FIRST_ROW, COL_COUNT) _
    ).Font.Bold = True
  rowCount = TABLE_FIRST_ROW + 1

  ' Loop through all sheets to find Validation
  For Each ws In ThisWorkbook.Worksheets
    ' Exclude the reference report worksheets to avoid duplications.
    If IsWsRefRpt(ws.codeName) Then GoTo Continue_ws
    
    Set validationRange = Nothing
    ' SpecialCells will error if NO cells have validation, so we keep this one
    On Error Resume Next
    Set validationRange = ws.Cells.SpecialCells(xlCellTypeAllValidation)
    err.Clear
    On Error GoTo Err_Proc
    
    If Not validationRange Is Nothing Then
      For Each targetCell In validationRange.Cells
        vType = XLDVTYPE_UNDEFINED
        On Error Resume Next
        vType = targetCell.Validation.Type
        err.Clear
        On Error GoTo Err_Proc
        
        'xlValidateList = 3, xlValidateCustom = 7;
        If vType = xlValidateList Or vType = xlValidateCustom Then
          rptWs.Cells(rowCount, 1).value = "'" & targetCell.Validation.Formula1
          rptWs.Cells(rowCount, 2).value = "'" & ws.name
          rptWs.Cells(rowCount, 3).value = "'" & targetCell.Address
          rptWs.Cells(rowCount, 4).value = "'" & "List"
          ' Prepend ' to ensure formulas are treated as text
          rowCount = rowCount + 1
        End If
      Next targetCell
    End If
Continue_ws:
  Next ws
  
  ' Convert Data Range into an Excel Table (ListObject)
  If rowCount > TABLE_FIRST_ROW + 2 Then
    Set tblRange = rptWs.Range( _
      rptWs.Cells(TABLE_FIRST_ROW, 1), _
      rptWs.Cells(rowCount - 1, COL_COUNT))
    Set rptTable = rptWs.ListObjects.Add( _
      SourceType:=xlSrcRange, _
      Source:=tblRange, _
      XlListObjectHasHeaders:=xlYes)
      
    With rptTable
      .name = TABLE_NAME
      
      With .Range.Font
        .name = "Aptos Narrow"
        .Size = 10
      End With ' .Range.Font
  
    ' Sorting Logic using the newly created Table
      With .Sort
      
        With .SortFields
          .Clear
          .Add2 key:=rptTable.ListColumns(COL_NAME_FORMULA_SOURCE).Range
          .Add2 key:=rptTable.ListColumns(COL_NAME_SHEET_NAME).Range
          .Add2 key:=rptTable.ListColumns(COL_NAME_CELL_ADDRESS).Range
        End With ' .SortFields
        
        .Header = xlYes
        .Apply
        
      End With ' .Sort
    End With ' . rptTable

  End If
    
  If creatingNewWs Then
    ' Auto-fit columns for clean presentation
    rptTable.Range.Columns.AutoFit
    rptWs.Protect
  Else
    ps.Restore rptWs
  End If
  
Exit_Proc:
  OptimizeAppEnvForSpeed False
  Const MSGBOX_TITLE = "Data Validation Audit Report"
  If errorOccurred Then
    ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
    MsgBox "Audit Failed.", vbCritical, MSGBOX_TITLE
  ElseIf Not silent Then
    MsgBox _
      "Audit Complete. Found " & (rowCount - 2) & " validation rules.", _
      vbInformation, _
      MSGBOX_TITLE
  End If
  Exit Sub
Err_Proc:
  errorOccurred = True
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Updates the defined names all references worksheet.
' Remarks: This sheet includes defined names referenced by:
'   1) Worksheet cells,
'   2, Lambda functions.
'   3) Data validaton sources, and
'   4) VBA code.
' Parameter(s)
'   fastMode - If True, the source worksheets are not updated; otherwise they
'     are updated with can be very slow. It is an optional parameter with
'     a default value of False.
' Date Created: 2026-06-14
' Date Last Modified: 2026-08-18
'------------------------------------------------------------------------------'
Public Sub UpsertDefNamesAllRefsRpt(Optional fastMode As Boolean = False)
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "UpsertDefNamesAllRefsRpt"
  
  Const BUTTON_UPDATE_FAST_NAME As String = "btnUpsrtDfNmsAllRfsRptFast"
  Const BUTTON_UPDATE_SLOW_NAME As String = "btnUpsrtDfNmsAllRfsRptSlow"

  Const DEF_NAME_PRINT_TITLES As String = "Print_Titles"
  Const DEST_TABLE_NAME As String = "tbl_DefNamesAllRefsRpt"
  Const DEST_TABLE_HEADER_ROW As Long = 5
  Const DEST_TABLE_COL_COUNT = 10
  Const DEST_COL_NAME_HEADER As String = "Name"
  Const DEST_COL_RWs_COUNT_HEADER As String = "RWs_Count"
  Const DEST_COL_RWs_NAMES_HEADER As String = "RWs_Names"
  Const DEST_COL_RF_COUNT_HEADER As String = "RF_Count"
  Const DEST_COL_RF_NAMES_HEADER As String = "RF_Names"
  Const DEST_COL_RDV_COUNT_HEADER As String = "RDV_Count"
  Const DEST_COL_RDV_NAMES_HEADER As String = "RDV_Names"
  Const DEST_COL_RVBA_COUNT_HEADER As String = "RVBA_Count"
  Const DEST_COL_RVBA_NAMES_HEADER As String = "RVBA_Names"
  Const DEST_COL_ANY_REFS_HEADER As String = "Any_Refs"
  Const DEST_COL_NAME As Long = 1 ' A
  Const DEST_COL_RWs_COUNT As Long = 2 ' B
  Const DEST_COL_RWs_NAMES As Long = 3 ' C
  Const DEST_COL_RF_COUNT As Long = 4 ' D
  Const DEST_COL_RF_NAMES As Long = 5 ' E
  Const DEST_COL_RDV_COUNT As Long = 6 ' F
  Const DEST_COL_RDV_NAMES As Long = 7 ' G
  Const DEST_COL_RVBA_COUNT As Long = 8 ' H
  Const DEST_COL_RVBA_NAMES As Long = 9 ' I
  Const DEST_COL_ANY_REFS As Long = 10 ' J
  Const DEST_WS_CODENAME As String = "SheetDefNamesAllRefs"
  Const DEST_WS_NAME As String = "Def Names All Refs"
  Const DEST_WS_TITLE = "Defined Names, All References"
  Const REF_COUNT_LIMIT = 10
  Const WS_REF_RPT_COL_HEADER_NAMES As String = "Name"
  Const WS_REF_RPT_COL_HEADER_REF_CNT As String = "Ref_Cnt"
  Const WS_REF_RPT_COL_HEADER_REFERENCES As String = "References"
  Const WS_REF_RPT_COL_HEADER_REFERS_TO As String = "Refers_To"
  Const WS_REF_RPT_COL_HEADER_VISIBLE As String = "Visible"
  Const WS_REF_RPT_TABLE_NAME As String = "tbl_DefNamesWsRefsRpt"
  Const TABLE_HEADER_BACKGROUND_COLORINDEX As Long = 8544277
  Const TABLE_HEADER_FOREGROUND_COLORINDEX As Long = 16777215
  Const VAL_RPT_HEADER_F1 = "F1"
  Const VAL_RPT_HEADER_F2 = "F2"
  Const VAL_RPT_HEADER_REF_COUNT = "Ref_Count"
  Const VAL_RPT_HEADER_REFERENCES = "References"
  Const VAL_RPT_TABLE_NAME As String = "tbl_DvF_Rpt"
  
  Dim creatingNewDestWs As Boolean
  Dim defName As String
  Dim defNameDestIdx As Long
  Dim defNameSrcIdx As Long
  Dim destRow As Long
  Dim destTbl As ListObject
  Dim destTblRng As Range
  Dim destTblRngStr As String
  Dim destUpdateButtonFast As Button
  Dim destUpdateButtonSlow As Button
  Dim destWs As Worksheet
  Dim rngCollection As Collection
  Dim errorOccurred As Boolean
  Dim formulaCell As Range
  Dim isInF1 As Boolean
  Dim isInF2 As Boolean
  Dim namesCount As Long
  Dim ps As clsProtectionState
  Dim refCount As Long
  Dim refCountLimitReached As Boolean
  Dim refCountStr As String
  Dim refCountAllNotZero As Boolean
  Dim refList As String
  Dim refListCollection As Collection
  Dim rng As Range
  Dim valTblRowCount As Long
  Dim wsRefRptDefNamesTbl As ListObject
  Dim wsRefRptMustBeUpdated As Boolean
  Dim wsRefRptNamesColRng As Range
  Dim wsRefRptRefCntColRng As Range
  Dim wsRefRptReferencesColRng As Range
  Dim wsRefRptRefersToColRng As Range
  Dim wsRefRptVisibleColRng As Range
  Dim valRptTbl As ListObject
  Dim valRptColF1Rng As Range
  Dim valRptColF2Rng As Range
  Dim valRptColRefCountRng As Range
  Dim valRptColReferencesRng As Range
  Dim vbComp As VBComponent

  errorOccurred = False
  
  On Error Resume Next
    
  ' Confirm that the Defined Names Worksheet References Report contains
  ' the information needed to update the Defined Names All References report.
  Set wsRefRptDefNamesTbl = Range(WS_REF_RPT_TABLE_NAME).ListObject
  If wsRefRptDefNamesTbl Is Nothing Then
    wsRefRptMustBeUpdated = True
    GoTo Exit_Proc
  End If
  
  On Error Resume Next
  Set wsRefRptRefCntColRng = _
    wsRefRptDefNamesTbl.ListColumns(WS_REF_RPT_COL_HEADER_REF_CNT).DataBodyRange
  err.Clear
  If wsRefRptRefCntColRng Is Nothing Then
    wsRefRptMustBeUpdated = True
    GoTo Exit_Proc
  End If
  
  On Error Resume Next
  Set wsRefRptReferencesColRng = _
    wsRefRptDefNamesTbl.ListColumns(WS_REF_RPT_COL_HEADER_REFERENCES).DataBodyRange
  err.Clear
  If wsRefRptReferencesColRng Is Nothing Then
    wsRefRptMustBeUpdated = True
    GoTo Exit_Proc
  End If
        
  If Not fastMode Then
    ' This takes a long time to run.
    UpsertDefNamesWsRefRpt fastMode:=False, silent:=True
  End If
  UpsertDvByFormulaRpt silent:=True

  OptimizeAppEnvForSpeed True
  
  Set destWs = GetWsByCodeName(DEST_WS_CODENAME)
  creatingNewDestWs = (destWs Is Nothing)
  
  If creatingNewDestWs Then
    Set destWs = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.count - 1))
    ThisWorkbook.VBProject.VBComponents(destWs.codeName).name = _
      DEST_WS_CODENAME
      
    With destWs
      .name = DEST_WS_NAME
      
      With .Cells.Font
        .name = "Aptos Narrow"
        .Size = 10
      End With
      
      With .Cells(1, 1)
        .value = DEST_WS_TITLE
        With .Font
          .name = "Aptos Display"
          .Size = 12
          .Bold = True
        End With
      End With
      
      ' Create the Headers that lie outside the formal table.
      Set rngCollection = New Collection
      
      .Range("B3").value = "Referer"
      rngCollection.Add .Range("B3:J3")
      
      .Range("B4").value = "Worksheets"
      rngCollection.Add .Range("B4:C4")
      
      .Range("D4").value = "Formulas (Including Lambda)"
      rngCollection.Add .Range("D4:E4")
      
      .Range("F4").value = "Data Validation Sources"
      rngCollection.Add .Range("F4:G4")
      
      .Range("H4").value = "VBA Modules"
      rngCollection.Add .Range("H4:I4")
      
      .Range("J4").value = "All"
      rngCollection.Add .Range("J4")
      
      For Each rng In rngCollection
        rng.Font.Color = TABLE_HEADER_FOREGROUND_COLORINDEX
        rng.Font.Bold = True
        rng.Interior.Color = TABLE_HEADER_BACKGROUND_COLORINDEX
        rng.HorizontalAlignment = xlCenterAcrossSelection
      Next rng

    End With
      
    ' Create Buttons
    Set destUpdateButtonSlow = destWs.buttons.Add( _
      Left:=300, _
      Top:=5, _
      Width:=200, _
      Height:=20)
    DoEvents ' Brief pause to let Excel register the object
    With destUpdateButtonSlow
      .OnAction = "RunUpsertDefNamesAllRptSlow"
      .Caption = "Complete Update (slow)"
      .name = BUTTON_UPDATE_SLOW_NAME
      .Placement = xlFreeFloating
    End With
    
    Set destUpdateButtonFast = destWs.buttons.Add( _
      Left:=300 + 225, _
      Top:=5, _
      Width:=200, _
      Height:=20)
    DoEvents ' Brief pause to let Excel register the object
    With destUpdateButtonFast
      .OnAction = "RunUpsertDefNamesAllRptFast"
      .Caption = "Fast Update (Sources Not Updated)"
      .name = BUTTON_UPDATE_FAST_NAME
      .Placement = xlFreeFloating
    End With
    
  Else
    Set ps = clsProtectionStateStatic.Create(destWs)
    destWs.Unprotect
    
    ' Clear existing ListObject if it exists to avoid conflicts
    On Error Resume Next
    destWs.ListObjects(DEST_TABLE_NAME).Delete
    err.Clear
    On Error GoTo Err_Proc
  End If
  
  With destWs
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_NAME).value = _
      DEST_COL_NAME_HEADER
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_RWs_COUNT).value = _
      DEST_COL_RWs_COUNT_HEADER
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_RWs_NAMES).value = _
      DEST_COL_RWs_NAMES_HEADER
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_RF_COUNT).value = _
      DEST_COL_RF_COUNT_HEADER
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_RF_NAMES).value = _
      DEST_COL_RF_NAMES_HEADER
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_RDV_COUNT).value = _
      DEST_COL_RDV_COUNT_HEADER
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_RDV_NAMES).value = _
      DEST_COL_RDV_NAMES_HEADER
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_RVBA_COUNT).value = _
      DEST_COL_RVBA_COUNT_HEADER
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_RVBA_NAMES).value = _
      DEST_COL_RVBA_NAMES_HEADER
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_ANY_REFS).value = _
      DEST_COL_ANY_REFS_HEADER
    .Range( _
      .Cells(DEST_TABLE_HEADER_ROW, 1), _
      .Cells(DEST_TABLE_HEADER_ROW, DEST_TABLE_COL_COUNT) _
      ).Font.Bold = True
  End With
  
  Set wsRefRptDefNamesTbl = Range(WS_REF_RPT_TABLE_NAME).ListObject
  Set wsRefRptNamesColRng = _
    wsRefRptDefNamesTbl.ListColumns(WS_REF_RPT_COL_HEADER_NAMES).DataBodyRange
  Set wsRefRptRefCntColRng = _
    wsRefRptDefNamesTbl.ListColumns(WS_REF_RPT_COL_HEADER_REF_CNT).DataBodyRange
  Set wsRefRptReferencesColRng = _
    wsRefRptDefNamesTbl.ListColumns(WS_REF_RPT_COL_HEADER_REFERENCES).DataBodyRange
  Set wsRefRptRefersToColRng = _
    wsRefRptDefNamesTbl.ListColumns(WS_REF_RPT_COL_HEADER_REFERS_TO).DataBodyRange
  Set wsRefRptVisibleColRng = _
    wsRefRptDefNamesTbl.ListColumns(WS_REF_RPT_COL_HEADER_VISIBLE).DataBodyRange
    
  Set valRptTbl = Range(VAL_RPT_TABLE_NAME).ListObject
  Set valRptColF1Rng = _
    valRptTbl.ListColumns(VAL_RPT_HEADER_F1).DataBodyRange
  Set valRptColF2Rng = _
    valRptTbl.ListColumns(VAL_RPT_HEADER_F2).DataBodyRange
  Set valRptColRefCountRng = _
    valRptTbl.ListColumns(VAL_RPT_HEADER_REF_COUNT).DataBodyRange
  Set valRptColReferencesRng = _
    valRptTbl.ListColumns(VAL_RPT_HEADER_REFERENCES).DataBodyRange
    
  namesCount = 0
  For defNameDestIdx = 1 To wsRefRptNamesColRng.count
    defName = wsRefRptNamesColRng(defNameDestIdx)
    If wsRefRptVisibleColRng(defNameDestIdx).value Then
      refCountAllNotZero = False
      namesCount = namesCount + 1
      destRow = DEST_TABLE_HEADER_ROW + namesCount
      destWs.Cells(destRow, DEST_COL_NAME).value = "'" & defName
      
      ' Print titles show up in the Define Names Worksheet References Report
      ' as having no references so we account for them here.
      ' They have format of 'WorksheetName'!Print_Titles
      If Right(defName, Len(DEF_NAME_PRINT_TITLES)) = DEF_NAME_PRINT_TITLES Then
        destWs.Cells(destRow, DEST_COL_RWs_COUNT).value = 1
        destWs.Cells(destRow, DEST_COL_RWs_NAMES).value = defName
        refCountAllNotZero = True
        GoTo Skip_For_Print_Title
      End If
      
      destWs.Cells(destRow, DEST_COL_RWs_COUNT).value = _
        "'" & wsRefRptRefCntColRng(defNameDestIdx)
      refCountAllNotZero = refCountAllNotZero Or (wsRefRptRefCntColRng(defNameDestIdx) <> 0)
      
      destWs.Cells(destRow, DEST_COL_RWs_NAMES).value = _
        "'" & wsRefRptReferencesColRng(defNameDestIdx)
      
      ' Get the formulas that reference the defined name.
      refCount = 0
      refList = ""
      refCountLimitReached = False
      For defNameSrcIdx = 1 To wsRefRptNamesColRng.count
        If defNameSrcIdx <> defNameDestIdx Then
          If ContainsString( _
            searchIn:=wsRefRptRefersToColRng(defNameSrcIdx), _
            searchFor:=defName, _
            wholeWordOnly:=True, _
            matchCase:=False _
            ) Then
            
            refCount = refCount + 1
            
            If refCount > REF_COUNT_LIMIT Then
              refCountLimitReached = True
              refList = refList + ", ..."
              Exit For
            End If
            
            If refCount <> 1 Then
              refList = refList & ", "
            End If
            
            refList = refList & "{" & wsRefRptRefersToColRng(defNameSrcIdx) & "}"

          End If
        End If
      Next defNameSrcIdx
      refCountAllNotZero = refCountAllNotZero Or (refCount > 0)
      refCountStr = _
        "'" & _
        IIf(refCountLimitReached, _
          ">" & REF_COUNT_LIMIT, _
          refCount)
      destWs.Cells(destRow, DEST_COL_RF_COUNT).value = refCountStr
      destWs.Cells(destRow, DEST_COL_RF_NAMES).value = "'" & refList
      
      ' Get the data validaton formulas that reference the defined name.
      valTblRowCount = valRptTbl.DataBodyRange.Rows.count
      refCount = 0
      Set refListCollection = New Collection
      For defNameSrcIdx = 1 To valTblRowCount
        isInF1 = ContainsString( _
          searchIn:=valRptColF1Rng(defNameSrcIdx), _
          searchFor:=defName, _
          wholeWordOnly:=True, _
          matchCase:=False _
          )
        isInF2 = ContainsString( _
          searchIn:=valRptColF2Rng(defNameSrcIdx), _
          searchFor:=defName, _
          wholeWordOnly:=True, _
          matchCase:=False _
          )
        If isInF1 Or isInF2 Then
          refCount = refCount + valRptColRefCountRng(defNameSrcIdx)
          refListCollection.Add valRptColReferencesRng(defNameSrcIdx)
        End If
      Next defNameSrcIdx
      refCountAllNotZero = refCountAllNotZero Or (refCount > 0)
      destWs.Cells(destRow, DEST_COL_RDV_COUNT).value = refCount
      Dim refListStr As String
      If refCount > 0 Then
        refListStr = CombineAndTruncateRefs( _
          refStrings:=refListCollection, _
          refCountLimit:=REF_COUNT_LIMIT)
      Else
        refListStr = ""
      End If
      
      destWs.Cells(destRow, DEST_COL_RDV_NAMES).value = "'" & refListStr
      
      ' Get the modules that reference the defined name.
      refCount = 0
      refList = ""
      refCountLimitReached = False
      For Each vbComp In ActiveWorkbook.VBProject.VBComponents
        If ComponentContainsString( _
          searchIn:=vbComp, _
          searchFor:=defName, _
          wholeWordOnly:=True, _
          matchCase:=False _
          ) Then
          
          refCount = refCount + 1
          
          If refCount > REF_COUNT_LIMIT Then
            refCountLimitReached = True
            refList = refList + ", ..."
            Exit For
          End If
          
          If refCount <> 1 Then
            refList = refList & ", "
          End If
          
          refList = refList & vbComp.name
          
        End If
      Next vbComp

      refCountStr = _
        "'" & _
        IIf(refCountLimitReached, _
          ">" & REF_COUNT_LIMIT, _
          refCount)
      destWs.Cells(destRow, DEST_COL_RVBA_COUNT).value = refCountStr
      destWs.Cells(destRow, DEST_COL_RVBA_NAMES).value = "'" & refList
      refCountAllNotZero = refCountAllNotZero Or (refCount > 0)
      
Skip_For_Print_Title:
      destWs.Cells(destRow, DEST_COL_ANY_REFS).value = refCountAllNotZero
    End If
  Next defNameDestIdx

  ' Convert the data range into an Excel Table (ListObject)
  If namesCount = 0 Then
    ' Create the table with a headers row on one blank data row to
    ' avoid an error withthe ListObjects.Add()
    destTblRngStr = "A5:J5"
  Else
    destTblRngStr = "A5:J" & namesCount + 5
  End If
  Set destTblRng = destWs.Range(destTblRngStr)
  Set destTbl = destWs.ListObjects.Add( _
    SourceType:=xlSrcRange, _
    Source:=destTblRng, _
    XlListObjectHasHeaders:=xlYes)
  With destTbl
    .name = DEST_TABLE_NAME
    
    With .Range.Font
      .name = "Aptos Narrow"
      .Size = 10
    End With
    
    With .Sort
      With .SortFields
        .Clear
        .Add2 key:=destTbl.ListColumns(DEST_COL_NAME).Range
      End With ' SortFields
      
      .Header = xlYes
      .Apply
    End With ' Sort
  End With ' destTbl
    
  If Not creatingNewDestWs Then
    ps.Restore destWs
  End If
  
Exit_Proc:
  OptimizeAppEnvForSpeed False
  Const MSGBOX_TITLE = "Defined Names, All References Report"
  If errorOccurred Then
    ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
    MsgBox "Update Failed.", vbCritical, MSGBOX_TITLE
  ElseIf wsRefRptMustBeUpdated Then
    Dim mbTitle
    mbTitle = _
      "The Defined Names Worksheet Reference Report " & vbCrLf & _
      "must be created/updated." & vbCrLf & vbCrLf & _
      "You must run a Complete Update to update that report " & vbCrLf & _
      "so that the Defined Names ALL References Report can be updated."
    MsgBox mbTitle, vbInformation, MSGBOX_TITLE
  Else
    MsgBox _
      "Update Complete.", vbInformation, MSGBOX_TITLE
  End If
  Exit Sub
Err_Proc:
  errorOccurred = True
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Creates a worksheet that list all defined names in this workbook
'   their properites: 1) Name, 2) RefersTo (definition), 3) Scope, 4) Value,
'   5) Visible, 6) Comment, 7) Count of references, 8) References.
' Parameter(s)
'   fastMode - if TRUE, dramatically speeds up method by not determining references
'     or reference counts; otherise, much slower and displays references and
'     reference counts. It is an optional parameter with a default value of
'     false.
'   silent - If True, the method does not display any messages to the user
'     except for errors; otherwise a report updated message is displayed to
'     the user when complete. It is an optional parameter with a default value
'     of False.
' Date Created: 2025-10-19
' Date Last Modified: 2026-08-18
'------------------------------------------------------------------------------'
Public Sub UpsertDefNamesWsRefRpt( _
  Optional fastMode As Boolean = False, _
  Optional silent As Boolean = False)
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "UpsertDefNamesWsRefRpt"
  
  Const FAST_BUTTON_NAME = "btnRunUpsrtDfNmsWsRptFastVrbs"
  Const SLOW_BUTTON_NAME = "btnRunUpsrtDfNmsWsRptSlowVrbs"
  
  ' report column numbers
  Const COL_COUNT = 8
  Const COL_HEADER_NAME As String = "Name"
  Const COL_HEADER_REFERS_TO As String = "Refers_To"
  Const COL_HEADER_SCOPE As String = "Scope"
  Const COL_HEADER_VALUE As String = "Value"
  Const COL_HEADER_VISIBLE As String = "Visible"
  Const COL_HEADER_COMMENT As String = "Comment"
  Const COL_HEADER_REF_COUNT As String = "Ref_Cnt"
  Const COL_HEADER_REFS As String = "References"
  Const COL_NUM_NAME As Integer = 1
  Const COL_NUM_REFERS_TO As Integer = 2
  Const COL_NUM_SCOPE As Integer = 3
  Const COL_NUM_VALUE As Integer = 4
  Const COL_NUM_VISIBLE As Integer = 5
  Const COL_NUM_COMMENT As Integer = 6
  Const COL_NUM_REF_COUNT As Integer = 7
  Const COL_NUM_REFS As Integer = 8

  
  Const REF_COUNT_LIMIT = 10
  Const REPORT_WORKSHEET_CODENAME As String = "SheetDefNamesWsRefsRpt"
  Const REPORT_WORKSHEET_NAME As String = "Def Names Ws Refs Rpt"
  Const TABLE_FIRST_ROW = 3
  Const TABLE_NAME As String = "tbl_DefNamesWsRefsRpt"
  Const REPORT_WORKSHEET_TITLE = "Defined Names, Worksheet References"
  
  Dim aCell As Range
  Dim defName As name

  Dim creatingNewWs As Boolean
  Dim errorOccurred As Boolean
  Dim firstRef As Boolean
  Dim inStrResult As Integer
  Dim mbButtons As VbMsgBoxStyle
  Dim mbPrompt As String
  Dim mbTitle As String
  Dim ps As clsProtectionState
  Dim nameValue As Variant

  Dim rptWs As Worksheet
  Dim refCount As Long
  Dim refCountLimitReached As Boolean
  Dim refCountStr As String
  Dim refList As String
  Dim rowNum As Long
  Dim rptTable As ListObject
  Dim tblRange As Range
  Dim testName As name

  Dim updateButtonFast As Button
  Dim updateButtonSlow As Button
  Dim wb As Workbook
  Dim ws As Worksheet
  Dim wsContainsRefs As Boolean
  errorOccurred = False
  
  OptimizeAppEnvForSpeed True
  
  Set wb = ThisWorkbook
  
  ' Create or clear the report worksheet
  Set rptWs = GetWsByCodeName(REPORT_WORKSHEET_CODENAME)
  creatingNewWs = (rptWs Is Nothing)
  
  If creatingNewWs Then
    Set rptWs = wb.Worksheets.Add(After:=Sheets(Sheets.count - 1))
    rptWs.name = REPORT_WORKSHEET_NAME
    ' This requires "Trust access to the VBA project object model" to be enabled and
    ' also for the workbook file to unblocked.
    ThisWorkbook.VBProject.VBComponents(rptWs.codeName).name = REPORT_WORKSHEET_CODENAME

    With rptWs.Cells.Font
        .name = "Aptos Narrow"
        .Size = 10
    End With
      
    With rptWs.Cells(1, 1)
      .value = REPORT_WORKSHEET_TITLE
      With .Font
        .name = "Aptos Display"
        .Size = 12
        .Bold = True
      End With
    End With
    
    ' Create Buttons
    Set updateButtonSlow = rptWs.buttons.Add( _
      Left:=300, _
      Top:=5, _
      Width:=150, _
      Height:=20)
    DoEvents ' Brief pause to let Excel register the object
    With updateButtonSlow
      .OnAction = "RunUpsrtDfNmsWsRptSlowVrbs"
      .Caption = "Complete Update (slow)"
      .name = SLOW_BUTTON_NAME
      .Placement = xlFreeFloating
    End With
    
    Set updateButtonFast = rptWs.buttons.Add( _
      Left:=300 + 175, _
      Top:=5, _
      Width:=150, _
      Height:=20)
    DoEvents ' Brief pause to let Excel register the object
    With updateButtonFast
      .OnAction = "RunUpsrtDfNmsWsRptFastVrbs"
      .Caption = "Fast Update (No references)"
      .name = FAST_BUTTON_NAME
      .Placement = xlFreeFloating
    End With
  Else
    Set ps = clsProtectionStateStatic.Create(rptWs)
    rptWs.Unprotect
    
    ' Clear existing table
    On Error Resume Next
    rptWs.ListObjects(TABLE_NAME).Delete
    err.Clear
    On Error GoTo Err_Proc
  End If
  
  ' Add headers
  With rptWs
    .Cells(TABLE_FIRST_ROW, COL_NUM_NAME) = COL_HEADER_NAME
    .Cells(TABLE_FIRST_ROW, COL_NUM_REFERS_TO) = COL_HEADER_REFERS_TO
    .Cells(TABLE_FIRST_ROW, COL_NUM_SCOPE) = COL_HEADER_SCOPE
    .Cells(TABLE_FIRST_ROW, COL_NUM_VALUE) = COL_HEADER_VALUE
    .Cells(TABLE_FIRST_ROW, COL_NUM_VISIBLE) = COL_HEADER_VISIBLE
    .Cells(TABLE_FIRST_ROW, COL_NUM_COMMENT) = COL_HEADER_COMMENT
    .Cells(TABLE_FIRST_ROW, COL_NUM_REF_COUNT) = _
      IIf(fastMode, "", COL_HEADER_REF_COUNT)
    .Cells(TABLE_FIRST_ROW, COL_NUM_REFS) = IIf(fastMode, "", COL_HEADER_REFS)
    .Range( _
      .Cells(TABLE_FIRST_ROW, 1), _
      .Cells(TABLE_FIRST_ROW, COL_COUNT) _
    ).Font.Bold = True
  End With
  
  rowNum = TABLE_FIRST_ROW + 1
  
  ' Loop through all defined names
  For Each defName In wb.Names
    ' Limit to 5 names to speed up testing.
    #If DEBUG_MODE Then
      If rowNum > TABLE_FIRST_ROW + 5 Then Exit For
    #End If
  
' Uncomment this block to processing names that are defined at both
' the Workbook and Worksheet level.
    ' If this a Worksheet scoped Name?
'    If Not (defName.Parent Is wb) Then
'      ' Does the same name exist as the Workbook scoped level.
'      ' If it does, skip the Worksheet scoped level.
'      On Error Resume Next
'      Set testName = wb.Names(defName.Name)
'      err.Clear
'      On Error GoTo Err_Proc
'      If Not testName Is Nothing Then
'         GoTo Continue_WbDefName ' Skip this iteration and go to the next name
'      End If
'    End If
    
    ' Defined name.
    rptWs.Cells(rowNum, COL_NUM_NAME).value = "'" & defName.name
         
    ' Use RefersTo, the defined name definition.
    rptWs.Cells(rowNum, COL_NUM_REFERS_TO).value = "'" & defName.RefersTo

    ' Scope
    If defName.Parent Is wb Then
      rptWs.Cells(rowNum, COL_NUM_SCOPE).value = "'Workbook"
    Else
      rptWs.Cells(rowNum, COL_NUM_SCOPE).value = "'" & defName.Parent.name
    End If
      
    ' Value
    nameValue = Application.Evaluate(defName.RefersTo)
    If IsError(nameValue) Then
        rptWs.Cells(rowNum, COL_NUM_VALUE).value = "'Error or Invalid"
    ElseIf IsArray(nameValue) Then
      ' Handle ranges (arrays)
      Dim valueString As String
      valueString = ""
      Dim i As Long, j As Long
      If TypeName(nameValue) = "Variant()" Then
        For i = LBound(nameValue, 1) To UBound(nameValue, 1)
          For j = LBound(nameValue, 2) To UBound(nameValue, 2)
            If IsError(nameValue(i, j)) Then
              Dim errorString
              Select Case nameValue(i, j)
                  Case CVErr(xlErrNA):     errorString = "#NA()"
                  Case CVErr(xlErrValue):  errorString = "#VALUE!"
                  Case CVErr(xlErrDiv0):   errorString = "#DIV/0!"
                  Case CVErr(xlErrRef):    errorString = "#REF!"
                  Case CVErr(xlErrName):   errorString = "#NAME?"
                  Case CVErr(xlErrNum):    errorString = "#NUM!"
                  Case CVErr(xlErrNull):   errorString = "#NULL!"
                  Case Else:               errorString = "#UNKNOWN_ERROR"
              End Select
              valueString = valueString & errorString
            Else
              valueString = valueString & nameValue(i, j)
            End If
          Next j
        Next i
        ' Remove trailing comma and space
        If Len(valueString) > 0 Then
          valueString = Left(valueString, Len(valueString) - 2)
        End If
        rptWs.Cells(rowNum, COL_NUM_VALUE).value = "'" & valueString
      Else
        rptWs.Cells(rowNum, COL_NUM_VALUE).value = "'Array (Non-standard)"
      End If
    Else
      ' Handle single values or formulas
      rptWs.Cells(rowNum, COL_NUM_VALUE).value = "'" & nameValue
    End If
    
    ' Visible
    rptWs.Cells(rowNum, COL_NUM_VISIBLE) = defName.Visible
    
    ' Comment
    rptWs.Cells(rowNum, COL_NUM_COMMENT) = "'" & defName.Comment
    
    If fastMode Then GoTo Skip_Reference_Search
    
    ' Search for cells using the name in formulas, excluding
    ' "Defined Names Report"
    refCountLimitReached = False
    refCount = 0
    refList = ""
    
    ' Non-visible defined names are used by Excel for backward compatibilty.
    ' The do not have Worksheet cells do not reference them directly.
    If Not defName.Visible Then GoTo Continue_WbDefName
    
    For Each ws In wb.Worksheets
      ' Searching any of the reference report worksheets might result in
      ' duplicates.
      If IsWsRefRpt(ws.codeName) Then GoTo Continue_ws
      
      wsContainsRefs = False
      firstRef = True
            
      For Each aCell In ws.UsedRange
        If Not aCell.HasFormula Then GoTo Continue_aCell
        
        inStrResult = InStr(1, aCell.formula, defName.name, vbTextCompare)
        
        If inStrResult = 0 Then GoTo Continue_aCell
        
        refCount = refCount + 1
        
        If refCount > REF_COUNT_LIMIT Then
          refCountLimitReached = True
          Exit For
        End If
      
        If firstRef Then
          If refCount <> 1 Then
            refList = refList & ";"
          End If
          refList = refList & "'" & ws.name & "'!"
          firstRef = False
        ' Except for the 1st cell address, add the cell address separator, ",".
        ElseIf refCount <> 1 Then
          refList = refList & ","
        End If
        
        refList = refList & aCell.Address
        wsContainsRefs = True
Continue_aCell:
      Next aCell
      
Continue_ws:
    Next ws
    
    If refCountLimitReached Then
      ' Indicate that the reference count exceeded the limit.
      refCountStr = ">" & REF_COUNT_LIMIT
      ' Indicate that not all the references are shown.
      refList = refList & "..."
    Else
      refCountStr = refCount
    End If
    
    ' Prepend "'" to force Excel to treat this value as text so that
    ' it will sort properly.
    rptWs.Cells(rowNum, COL_NUM_REF_COUNT).value = "'" & refCountStr
    rptWs.Cells(rowNum, COL_NUM_REFS).value = "'" & refList
    
Skip_Reference_Search:
    rowNum = rowNum + 1

Continue_WbDefName:
  Next defName
  
  ' Convert Data Range into an Excel Table (ListObject)
  ' If no names were found, the table's first row will be empty
  With rptWs
    Set tblRange = .Range( _
      .Cells(TABLE_FIRST_ROW, 1), _
      .Cells(rowNum - 1, IIf(fastMode, COL_COUNT - 2, COL_COUNT)) _
      )
  End With
  
  Set rptTable = rptWs.ListObjects.Add( _
    SourceType:=xlSrcRange, _
    Source:=tblRange, _
    XlListObjectHasHeaders:=xlYes)
  
  With rptTable
    .name = TABLE_NAME
    
    With .Range.Font
      .name = "Aptos Narrow"
      .Size = 10
    End With
    
    With .Sort
    
      With .SortFields
        .Clear
        .Add2 key:=rptTable.ListColumns(COL_NUM_VISIBLE).Range
        .Add2 key:=rptTable.ListColumns(COL_NUM_NAME).Range
        .Add2 key:=rptTable.ListColumns(COL_NUM_SCOPE).Range
      End With 'SortFields
      
      .Header = xlYes
      .Apply
    End With ' Sort
  End With ' rptTable

  If creatingNewWs Then
    rptWs.Protect
  Else
    ps.Restore rptWs
  End If
  
Exit_Proc:
  OptimizeAppEnvForSpeed False
  mbTitle = REPORT_WORKSHEET_TITLE
  If errorOccurred Then
    ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
    mbPrompt = "Defined Names Report generation/update failed"
    mbButtons = vbCritical
  Else
    mbPrompt = "Defined Names Report generated/updated in the " & vbCrLf & _
      """" & REPORT_WORKSHEET_NAME & """ Worksheet."
    mbButtons = vbInformation
  End If
  
  If Not silent Then
    MsgBox prompt:=mbPrompt, buttons:=mbButtons, title:=mbTitle
  End If
  
  Exit Sub
Err_Proc:
  errorOccurred = True
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Creates or updates a worksheet that shows data validation
'   properties for all cells in the workbook that have data validation defined.
' Parameter(s)
'   silent - It True, the method does not display any messages to the user
'     except for errors; otherwise a report updated message is displayed to
'     the user when complete. It is an optional parameter with a default value
'     of False.'
' Date Created: 2026-07-07
' Date Last Modified: 2026-08-18
'------------------------------------------------------------------------------'
Public Sub UpsertDvAllRpt(Optional silent As Boolean = False)
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "UpsertDvAllRpt"

  Const BUTTON_NAME As String = "btnUpsertDvAllRpt"
  
  Const COL_IDX_SHEET As Long = 1
  Const COL_IDX_CELL As Long = COL_IDX_SHEET + 1 ' 2
  Const COL_IDX_VAL_TYPE As Long = COL_IDX_CELL + 1 ' 3
  Const COL_IDX_OPERATOR As Long = COL_IDX_VAL_TYPE + 1 ' 4
  Const COL_IDX_F1 As Long = COL_IDX_OPERATOR + 1 ' 5
  Const COL_IDX_F1_TYPE As Long = COL_IDX_F1 + 1 ' 6
  Const COL_IDX_F2 As Long = COL_IDX_F1_TYPE + 1 ' 7
  Const COL_IDX_F2_TYPE As Long = COL_IDX_F2 + 1 ' 8
  
  Const COL_COUNT As Long = COL_IDX_F2_TYPE

  Const COL_HDR_SHEET As String = "Sheet"
  Const COL_HDR_CELL As String = "Cell"
  Const COL_HDR_VAL_TYPE As String = "Val_Type"
  Const COL_HDR_OPERATOR As String = "Operator"
  Const COL_HDR_F1 As String = "F1"
  Const COL_HDR_F1_TYPE As String = "F1_Type"
  Const COL_HDR_F2 As String = "F2"
  Const COL_HDR_F2_TYPE As String = "F2_Type"
  
  Const RPT_WS_CODENAME As String = _
    "SheetDvAllRpt"
  Const RPT_WS_NAME = "DV All Rpt"
  Const RPT_WS_TITLE = "Data Validation, All Cells, Report"
  Const TBL_FIRST_ROW As Long = 3
  Const TBL_NAME As String = "tbl_DvAllRpt"

  Dim creatingNewWs As Boolean
  Dim errorOccurred As Boolean
  Dim f1 As String
  Dim f1Type As ValidationSrcType
  Dim f2 As String
  Dim f2Type As ValidationSrcType
  Dim op As XlFormatConditionOperator
  Dim ps As clsProtectionState
  Dim rowCount As Long
  Dim rptTable As ListObject
  Dim rptWs As Worksheet
  Dim targetCell As Range
  Dim tblRange As Range
  Dim updateButton As Button
  Dim valRng As Range
  Dim vType As XlDVType
  Dim ws As Worksheet
  Dim wb As Workbook

  errorOccurred = False

  OptimizeAppEnvForSpeed True

  Set wb = ThisWorkbook

  ' Create or clear the audit sheet
  Set rptWs = _
    GetWsByCodeName(RPT_WS_CODENAME)
  creatingNewWs = (rptWs Is Nothing)

  If creatingNewWs Then
    Set rptWs = wb.Worksheets.Add(After:=Sheets(Sheets.count - 1))
    ThisWorkbook.VBProject.VBComponents(rptWs.codeName).name = _
      RPT_WS_CODENAME

    With rptWs
      .name = RPT_WS_NAME

      With .Cells.Font
        .name = "Aptos Narrow"
        .Size = 10
      End With

      With .Cells(1, 1)
        .value = RPT_WS_TITLE
        With .Font
          .name = "Aptos Display"
          .Size = 12
          .Bold = True
        End With
      End With

    End With

    ' Create Button: Left, Top, Width, Height (Positioned near Column F)
    Set updateButton = rptWs.buttons.Add( _
      Left:=300, _
      Top:=5, _
      Width:=80, _
      Height:=20)
    DoEvents ' Brief pause to let Excel register the object
    With updateButton
      .OnAction = "RunUpsertDvAllRptVrbse"
      .Caption = "Update"
      .name = BUTTON_NAME
      .Placement = xlFreeFloating
    End With
  Else
    Set ps = clsProtectionStateStatic.Create(rptWs)
    rptWs.Unprotect
    ' Clear existing ListObject if it exists to avoid conflicts
    On Error Resume Next
    rptWs.ListObjects(TBL_NAME).Delete
    err.Clear
    On Error GoTo Err_Proc
  End If

  ' Headers
  rptWs.Cells(TBL_FIRST_ROW, COL_IDX_SHEET).value = _
    COL_HDR_SHEET ' Col 1
  rptWs.Cells(TBL_FIRST_ROW, COL_IDX_CELL).value = _
    COL_HDR_CELL ' Col 2
  rptWs.Cells(TBL_FIRST_ROW, COL_IDX_VAL_TYPE).value = _
    COL_HDR_VAL_TYPE ' Col 3
  rptWs.Cells(TBL_FIRST_ROW, COL_IDX_OPERATOR).value = _
    COL_HDR_OPERATOR ' Col 4
  rptWs.Cells(TBL_FIRST_ROW, COL_IDX_F1).value = _
    COL_HDR_F1 ' Col 5
  rptWs.Cells(TBL_FIRST_ROW, COL_IDX_F1_TYPE).value = _
    COL_HDR_F1_TYPE ' Col 6
  rptWs.Cells(TBL_FIRST_ROW, COL_IDX_F2).value = _
    COL_HDR_F2 ' Col 7
  rptWs.Cells(TBL_FIRST_ROW, COL_IDX_F2_TYPE).value = _
    COL_HDR_F2_TYPE ' Col 8
    
  rptWs.Range( _
    rptWs.Cells(TBL_FIRST_ROW, 1), _
    rptWs.Cells(TBL_FIRST_ROW, COL_COUNT) _
    ).Font.Bold = True
  rowCount = TBL_FIRST_ROW + 1
  
  ' Loop through all sheets to find Validation
  For Each ws In ThisWorkbook.Worksheets
    ' Exclude the reference report worksheets to avoid duplications.
    If IsWsRefRpt(ws.codeName) Then GoTo Continue_ws
    
    Set valRng = Nothing
    ' An attempt reference to Range.SpecialCells(xlCellTypeAllValidation)
    ' will raise an error if no cells in the range have validation.
    On Error Resume Next
    Set valRng = ws.Cells.SpecialCells(xlCellTypeAllValidation)
    err.Clear
    On Error GoTo Err_Proc
    
    ' If no cells in this worksheet have defined data validation, then
    ' go to next worksheet iteration.
    If valRng Is Nothing Then GoTo Continue_ws
    
    For Each targetCell In valRng.Cells
      ' Gemini suggested error checking here but the Set valRng statement
      ' above should insure each cell in the range has validation defined.

      rptWs.Cells(rowCount, COL_IDX_SHEET).value = _
        "'" & targetCell.Parent.name ' Col 1
      rptWs.Cells(rowCount, COL_IDX_CELL).value = _
        "'" & targetCell.Address ' Col 2
      rptWs.Cells(rowCount, COL_IDX_VAL_TYPE).value = _
        "'" & XlDVTypeToString(targetCell.Validation.Type) ' Col 3
      GetDvFormulasAndTypes _
        targetCell:=targetCell, _
        op:=op, _
        f1:=f1, _
        f1Type:=f1Type, _
        f2:=f2, _
        f2Type:=f2Type
      rptWs.Cells(rowCount, COL_IDX_OPERATOR).value = _
        "'" & XlFormatConditionOperatorToString(op) ' col 4
      rptWs.Cells(rowCount, COL_IDX_F1).value = "'" & f1 ' col 5
      rptWs.Cells(rowCount, COL_IDX_F1_TYPE).value = _
        "'" & ValSrcTypeToString(f1Type) ' Col 6
      rptWs.Cells(rowCount, COL_IDX_F2).value = "'" & f2 ' col 7
      rptWs.Cells(rowCount, COL_IDX_F2_TYPE).value = _
        "'" & ValSrcTypeToString(f2Type) ' Col 8
      rowCount = rowCount + 1
    Next targetCell
    
Continue_ws:
  Next ws

  ' Convert Data Range into an Excel Table (ListObject)
  If rowCount > TBL_FIRST_ROW + 2 Then
    Set tblRange = rptWs.Range( _
      rptWs.Cells(TBL_FIRST_ROW, 1), _
      rptWs.Cells(rowCount - 1, COL_COUNT))
  ' If no data validation was found, create table with header
  ' and one blank row.
  Else
    Set tblRange = rptWs.Range( _
      rptWs.Cells(TBL_FIRST_ROW, 1), _
      rptWs.Cells(TBL_FIRST_ROW + 1, COL_COUNT))
  End If
  
  Set rptTable = rptWs.ListObjects.Add( _
    SourceType:=xlSrcRange, _
    Source:=tblRange, _
    XlListObjectHasHeaders:=xlYes)

  With rptTable
    .name = TBL_NAME

    With .Range.Font
      .name = "Aptos Narrow"
      .Size = 10
    End With ' .Range.Font

  ' Sorting Logic using the newly created Table
    With .Sort

      With .SortFields
        .Clear
        .Add2 key:=rptTable.ListColumns(COL_IDX_SHEET).Range
        .Add2 key:=rptTable.ListColumns(COL_IDX_CELL).Range
      End With ' .SortFields

      .Header = xlYes
      .Apply

    End With ' .Sort
  End With ' . rptTable

  If creatingNewWs Then
    ' Auto-fit columns for clean presentation
    rptTable.Range.Columns.AutoFit
    rptWs.Protect
  Else
    ps.Restore rptWs
  End If

Exit_Proc:
  OptimizeAppEnvForSpeed False
  Const MSGBOX_TITLE = RPT_WS_TITLE
  If errorOccurred Then
    ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
    MsgBox "There was a problem creating the report.", vbCritical, MSGBOX_TITLE
  ElseIf Not silent Then
    MsgBox _
      "Report complete. " & _
        "Found " & (rowCount - 2) & " cells with data validation.", _
      vbInformation, _
      MSGBOX_TITLE
  End If
  Exit Sub
Err_Proc:
  errorOccurred = True
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Creates or updates a worksheet that shows data validation
'   formulas.
' Parameter(s)
'   silent - It True, the method does not display any messages to the user
'     except for errors; otherwise a report updated message is displayed to
'     the user when complete. It is an optional parameter with a default value
'     of False.'
' Date Created: 2026-07-07
' Date Last Modified: 2026-08-18
'------------------------------------------------------------------------------'
Public Sub UpsertDvByFormulaRpt(Optional silent As Boolean = False)
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "UpsertDvByFormulaRpt"

  Const BUTTON_NAME As String = "btnUpsertDvByFormulaRpt"
  
  Const COL_IDX_F1 As Long = 1
  Const COL_IDX_F2 As Long = COL_IDX_F1 + 1 ' 2
  Const COL_IDX_REF_COUNT As Long = COL_IDX_F2 + 1 ' 3
  Const COL_IDX_REFS As Long = COL_IDX_REF_COUNT + 1 ' 4
  
  Const COL_COUNT As Long = COL_IDX_REFS

  Const COL_HDR_F1 As String = "F1"
  Const COL_HDR_F2 As String = "F2"
  Const COL_HDR_REF_COUNT As String = "Ref_Count"
  Const COL_HDR_REFS As String = "References"
  
  Const REF_COUNT_LIMIT As Long = 25
  
  Const RPT_WS_CODENAME As String = _
    "SheetDvByForumlaRpt"
  Const RPT_WS_NAME = "DV by F Rpt"
  Const RPT_WS_TITLE = "Formulas Used for Data Validation Report"
  Const TBL_FIRST_ROW As Long = 3
  Const TBL_NAME As String = "tbl_DvF_Rpt"

  Dim creatingNewWs As Boolean
  Dim errorOccurred As Boolean

  Dim formulaKey As String
  Dim formulaRefInfoDict As Dictionary
  Dim formulaRefInfo As clsDVFormulaRefInfo
  Dim ps As clsProtectionState
  Dim refCountLimitExceeded As Boolean
  Dim rowCount As Long
  Dim rptTable As ListObject
  Dim rptWs As Worksheet
  Dim targetCell As Range
  Dim tblRange As Range
  Dim updateButton As Button
  Dim valRng As Range
  Dim varFormulaKey As Variant
  Dim ws As Worksheet
  Dim wb As Workbook

  errorOccurred = False

  OptimizeAppEnvForSpeed True

  Set wb = ThisWorkbook

  ' Create or clear the audit sheet
  Set rptWs = _
    GetWsByCodeName(RPT_WS_CODENAME)
  creatingNewWs = (rptWs Is Nothing)

  If creatingNewWs Then
    Set rptWs = wb.Worksheets.Add(After:=Sheets(Sheets.count - 1))
    ThisWorkbook.VBProject.VBComponents(rptWs.codeName).name = _
      RPT_WS_CODENAME

    With rptWs
      .name = RPT_WS_NAME

      With .Cells.Font
        .name = "Aptos Narrow"
        .Size = 10
      End With

      With .Cells(1, 1)
        .value = RPT_WS_TITLE
        With .Font
          .name = "Aptos Display"
          .Size = 12
          .Bold = True
        End With
      End With

    End With

    ' Create Button: Left, Top, Width, Height (Positioned near Column F)
    Set updateButton = rptWs.buttons.Add( _
      Left:=300, _
      Top:=5, _
      Width:=80, _
      Height:=20)
    DoEvents ' Brief pause to let Excel register the object
    With updateButton
      .OnAction = "UpsertDvByFormulaRpt"
      .Caption = "Update"
      .name = BUTTON_NAME
      .Placement = xlFreeFloating
    End With
  Else
    Set ps = clsProtectionStateStatic.Create(rptWs)
    rptWs.Unprotect
    ' Clear existing ListObject if it exists to avoid conflicts
    On Error Resume Next
    rptWs.ListObjects(TBL_NAME).Delete
    err.Clear
    On Error GoTo Err_Proc
  End If

  ' Headers
  rptWs.Cells(TBL_FIRST_ROW, COL_IDX_F1).value = _
    COL_HDR_F1 ' Col 1
  rptWs.Cells(TBL_FIRST_ROW, COL_IDX_F2).value = _
    COL_HDR_F2 ' Col 2
  rptWs.Cells(TBL_FIRST_ROW, COL_IDX_REF_COUNT).value = _
    COL_HDR_REF_COUNT ' Col 3
  rptWs.Cells(TBL_FIRST_ROW, COL_IDX_REFS).value = _
    COL_HDR_REFS ' Col 4
    
  rptWs.Range( _
    rptWs.Cells(TBL_FIRST_ROW, 1), _
    rptWs.Cells(TBL_FIRST_ROW, COL_COUNT) _
    ).Font.Bold = True
  rowCount = TBL_FIRST_ROW + 1
  
  Set formulaRefInfoDict = New Dictionary
  ' Loop through all sheets to find Validation
  For Each ws In ThisWorkbook.Worksheets
    ' Exclude the reference report worksheets to avoid duplications.
    If IsWsRefRpt(ws.codeName) Then GoTo Continue_ws

    Set valRng = Nothing
    ' An attempt reference to Range.SpecialCells(xlCellTypeAllValidation)
    ' will raise an error if no cells in the range have validation.
    On Error Resume Next
    Set valRng = ws.Cells.SpecialCells(xlCellTypeAllValidation)
    err.Clear
    On Error GoTo Err_Proc
    
    ' If no cells in this worksheet have defined data validation, then
    ' go to next worksheet iteration.
    If valRng Is Nothing Then GoTo Continue_ws
    
    For Each targetCell In valRng.Cells
      ' Gemini suggested error checking here but the Set valRng statement
      ' above should insure each cell in the range has validation defined.
            
      Set formulaRefInfo = _
        clsDVFormulaRefInfoStatic.CreateFromCell(pCell:=targetCell)
      formulaKey = formulaRefInfo.formulaKey
      If Not formulaRefInfoDict.Exists(formulaKey) Then
        formulaRefInfoDict.Add key:=formulaKey, item:=formulaRefInfo
      End If
      Set formulaRefInfo = formulaRefInfoDict(formulaKey)
      
      formulaRefInfo.AddReference _
        wsName:=ws.name, _
        cellAddress:=targetCell.Address
    Next targetCell
Continue_ws:
  Next ws
      
  For Each varFormulaKey In formulaRefInfoDict.Keys
    Set formulaRefInfo = formulaRefInfoDict(varFormulaKey)
    rptWs.Cells(rowCount, COL_IDX_F1).value = _
      "'" & formulaRefInfo.Formula1 ' col 1
    rptWs.Cells(rowCount, COL_IDX_F2).value = _
      "'" & formulaRefInfo.formula2 ' col 2
    rptWs.Cells(rowCount, COL_IDX_REF_COUNT).value = _
      formulaRefInfo.GetRefCount ' col 3
    rptWs.Cells(rowCount, COL_IDX_REFS).value = _
      "'" & formulaRefInfo.GetRefString( _
        refLimitExceeded:=refCountLimitExceeded, _
        refLimit:=REF_COUNT_LIMIT) ' col 4
    rowCount = rowCount + 1
  Next varFormulaKey

  ' Convert Data Range into an Excel Table (ListObject)
  If rowCount > TBL_FIRST_ROW + 2 Then
    Set tblRange = rptWs.Range( _
      rptWs.Cells(TBL_FIRST_ROW, 1), _
      rptWs.Cells(rowCount - 1, COL_COUNT))
  ' If no data validation was found, create table with header
  ' and one blank row.
  Else
    Set tblRange = rptWs.Range( _
      rptWs.Cells(TBL_FIRST_ROW, 1), _
      rptWs.Cells(TBL_FIRST_ROW + 1, COL_COUNT))
  End If
  
  Set rptTable = rptWs.ListObjects.Add( _
    SourceType:=xlSrcRange, _
    Source:=tblRange, _
    XlListObjectHasHeaders:=xlYes)

  With rptTable
    .name = TBL_NAME

    With .Range.Font
      .name = "Aptos Narrow"
      .Size = 10
    End With ' .Range.Font

  ' Sorting Logic using the newly created Table
    With .Sort

      With .SortFields
        .Clear
        .Add2 key:=rptTable.ListColumns(COL_IDX_F1).Range
        .Add2 key:=rptTable.ListColumns(COL_IDX_F2).Range
      End With ' .SortFields

      .Header = xlYes
      .Apply

    End With ' .Sort
  End With ' . rptTable

  If creatingNewWs Then
    ' Auto-fit columns for clean presentation
    rptTable.Range.Columns.AutoFit
    rptWs.Protect
  Else
    ps.Restore rptWs
  End If

Exit_Proc:
  OptimizeAppEnvForSpeed False
  Const MSGBOX_TITLE = RPT_WS_TITLE
  If errorOccurred Then
    ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
    MsgBox _
     "There was a problem creating the report.", _
     vbCritical, _
     MSGBOX_TITLE
  ElseIf Not silent Then
    MsgBox _
      "Report complete. " & _
        "Found " & (rowCount - 4) & " unique formulas used for data validation.", _
      vbInformation, _
      MSGBOX_TITLE
  End If
  Exit Sub
Err_Proc:
  errorOccurred = True
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Creates or updates a worksheet that list all tables in this
'   workbook and their properites: 1) Name, 2) Address ), 3) Comment, 5) Count of
'   references by worksheet cells, and 6) A a list cell references.
' Parameter(s)
'   fastMode - if TRUE, dramatically speeds up method by not determining references
'     or reference counts; otherise, much slower and displays references and
'     reference counts. It is an optional parameter with a default value of
'     false.
'   silent - If True, the method does not display any messages to the user
'     except for errors; otherwise a report updated message is displayed to
'     the user when complete. It is an optional parameter with a default value
'     of False.
' Date Created: 2026-07-11
' Date Last Modified: 2026-08-18
'------------------------------------------------------------------------------'
Public Sub UpsertTblsWsRefRpt( _
  Optional fastMode As Boolean = False, _
  Optional silent As Boolean = False)
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "UpsertTblsWsRefRpt"
  
  Const FAST_BUTTON_NAME = "btnRunUpsrtTblsWsWsRptFastVrbs"
  Const SLOW_BUTTON_NAME = "btnRunUpsrtTblsWsWsRptSlowVrbs"
  
  ' report column numbers
  Const COL_HEADER_NAME As String = "Name"
  Const COL_HEADER_ADDRESS As String = "Address"
  Const COL_HEADER_COMMENT As String = "Comment"
  Const COL_HEADER_REF_COUNT As String = "Ref_Cnt"
  Const COL_HEADER_REFS As String = "References"
  Const COL_NUM_NAME As Integer = 1
  Const COL_NUM_ADDRESS As Integer = 2
  Const COL_NUM_COMMENT As Integer = 3
  Const COL_NUM_REF_COUNT As Integer = 4
  Const COL_NUM_REFS As Integer = 5
  Const COL_COUNT = 5

  Const REF_COUNT_LIMIT = 10
  Const REPORT_WORKSHEET_CODENAME As String = "SheetTblsWsRefsRpt"
  Const REPORT_WORKSHEET_NAME As String = "Tables Ws Refs Rpt"
  Const TABLE_FIRST_ROW = 3
  Const TABLE_NAME As String = "tbl_TblWsRefsRpt"
  Const REPORT_WORKSHEET_TITLE = "Tables, Worksheet References"
  
  Dim creatingNewWs As Boolean
  Dim errorOccurred As Boolean
  Dim firstRef As Boolean
  Dim inStrResult As Integer
  Dim mbButtons As VbMsgBoxStyle
  Dim mbPrompt As String
  Dim mbTitle As String
  Dim nameValue As Variant
  Dim ps As clsProtectionState
  Dim rptWs As Worksheet
  Dim refCount As Long
  Dim refCountLimitReached As Boolean
  Dim refCountStr As String
  Dim refList As String
  Dim rowNum As Long
  Dim rptTable As ListObject
  Dim targetCell As Range
  Dim targetTbl As ListObject
  Dim tblRange As Range
  Dim updateButtonFast As Button
  Dim updateButtonSlow As Button
  Dim wb As Workbook
  Dim wsForTblSearch As Worksheet
  Dim wsForRefSearch As Worksheet
  Dim wsContainsRefs As Boolean
  
  errorOccurred = False
  
  OptimizeAppEnvForSpeed True
  
  Set wb = ThisWorkbook
  
  ' Create or clear the report worksheet
  Set rptWs = GetWsByCodeName(REPORT_WORKSHEET_CODENAME)
  creatingNewWs = (rptWs Is Nothing)
  
  If creatingNewWs Then
    Set rptWs = wb.Worksheets.Add(After:=Sheets(Sheets.count - 1))
    rptWs.name = REPORT_WORKSHEET_NAME
    ' This requires "Trust access to the VBA project object model" to be enabled and
    ' also for the workbook file to unblocked.
    ThisWorkbook.VBProject.VBComponents(rptWs.codeName).name = REPORT_WORKSHEET_CODENAME

    With rptWs.Cells.Font
        .name = "Aptos Narrow"
        .Size = 10
    End With
      
    With rptWs.Cells(1, 1)
      .value = REPORT_WORKSHEET_TITLE
      With .Font
        .name = "Aptos Display"
        .Size = 12
        .Bold = True
      End With
    End With
    
    ' Create Buttons
    Set updateButtonSlow = rptWs.buttons.Add( _
      Left:=300, _
      Top:=5, _
      Width:=150, _
      Height:=20)
    DoEvents ' Brief pause to let Excel register the object
    With updateButtonSlow
      .OnAction = "RunUpsrtTblsWsRptSlowVrbs"
      .Caption = "Complete Update (slow)"
      .name = SLOW_BUTTON_NAME
      .Placement = xlFreeFloating
    End With
    
    Set updateButtonFast = rptWs.buttons.Add( _
      Left:=300 + 175, _
      Top:=5, _
      Width:=150, _
      Height:=20)
    DoEvents ' Brief pause to let Excel register the object
    With updateButtonFast
      .OnAction = "RunUpsrtTblsWsRptFastVrbs"
      .Caption = "Fast Update (No references)"
      .name = FAST_BUTTON_NAME
      .Placement = xlFreeFloating
    End With
  Else
    Set ps = clsProtectionStateStatic.Create(rptWs)
    rptWs.Unprotect
    
    ' Clear existing table
    On Error Resume Next
    rptWs.ListObjects(TABLE_NAME).Delete
    err.Clear
    On Error GoTo Err_Proc
  End If
  
  ' Add headers
  With rptWs
    .Cells(TABLE_FIRST_ROW, COL_NUM_NAME) = COL_HEADER_NAME
    .Cells(TABLE_FIRST_ROW, COL_NUM_ADDRESS) = COL_HEADER_ADDRESS
    .Cells(TABLE_FIRST_ROW, COL_NUM_COMMENT) = COL_HEADER_COMMENT
    .Cells(TABLE_FIRST_ROW, COL_NUM_REF_COUNT) = _
      IIf(fastMode, "", COL_HEADER_REF_COUNT)
    .Cells(TABLE_FIRST_ROW, COL_NUM_REFS) = IIf(fastMode, "", COL_HEADER_REFS)
    .Range( _
      .Cells(TABLE_FIRST_ROW, 1), _
      .Cells(TABLE_FIRST_ROW, COL_COUNT) _
    ).Font.Bold = True
  End With
  
  rowNum = TABLE_FIRST_ROW + 1
  
  ' Loop through all the tables. They have global scope but
  ' are contained by individual worksheets.
  For Each wsForTblSearch In wb.Worksheets
    ' Exclude the reference report worksheets to avoid duplications.
    If IsWsRefRpt(wsForTblSearch.codeName) Then GoTo Continue_wsForTblSearch
    
    If wsForTblSearch.ListObjects.count = 0 Then GoTo Continue_wsForTblSearch
    
    For Each targetTbl In wsForTblSearch.ListObjects
      ' Limit to 5 names to speed up testing.
      #If DEBUG_MODE Then
        If rowNum > TABLE_FIRST_ROW + 5 Then Exit For ' targetTbl
      #End If
     
      ' Table name.
      rptWs.Cells(rowNum, COL_NUM_NAME).value = "'" & targetTbl.name
           
      ' Table Range Address
      rptWs.Cells(rowNum, COL_NUM_ADDRESS).value = _
        "'" & "'" & wsForTblSearch.name & "'!" & targetTbl.Range.Address
        
      ' Comment
      rptWs.Cells(rowNum, COL_NUM_COMMENT) = "'" & targetTbl.Comment
      
      If fastMode Then GoTo Skip_Reference_Search
      
      ' Search for cells using the name in formulas, excluding
      ' the reference report tables
      refCountLimitReached = False
      refCount = 0
      refList = ""
      
      For Each wsForRefSearch In wb.Worksheets
        ' Exclude the reference report worksheets to avoid duplications.
        If IsWsRefRpt(wsForRefSearch.codeName) Then
          GoTo Continue_WsForRefSearch
        End If
        
        wsContainsRefs = False
        firstRef = True
        
        For Each targetCell In wsForRefSearch.UsedRange
          If Not targetCell.HasFormula Then GoTo Continue_targetCell
          
          inStrResult = InStr(1, targetCell.formula, targetTbl.name, vbTextCompare)
          If inStrResult = 0 Then GoTo Continue_targetCell
          
          refCount = refCount + 1
          
          If refCount > REF_COUNT_LIMIT Then
            refCountLimitReached = True
            Exit For
          End If
        
          If firstRef Then
            If refCount <> 1 Then
              refList = refList & ";"
            End If
            refList = refList & "'" & wsForRefSearch.name & "'!"
            firstRef = False
          ' Except for the 1st cell address, add the cell address separator, ",".
          ElseIf refCount <> 1 Then
            refList = refList & ","
          End If
          
          refList = refList & targetCell.Address
          wsContainsRefs = True
Continue_targetCell:
        Next targetCell
        
Continue_WsForRefSearch:
      Next wsForRefSearch
      
      If refCountLimitReached Then
        ' Indicate that the reference count exceeded the limit.
        refCountStr = ">" & REF_COUNT_LIMIT
        ' Indicate that not all the references are shown.
        refList = refList & "..."
      Else
        refCountStr = refCount
      End If
      
      ' Prepend "'" to force Excel to treat this value as text so that
      ' it will sort properly.
      rptWs.Cells(rowNum, COL_NUM_REF_COUNT).value = "'" & refCountStr
      rptWs.Cells(rowNum, COL_NUM_REFS).value = "'" & refList
      
Skip_Reference_Search:
      rowNum = rowNum + 1
  
Cotinue_targetTbl:
    Next targetTbl
    
Continue_wsForTblSearch:
  Next wsForTblSearch
  
  ' Convert Data Range into an Excel Table (ListObject)
  ' If no names were found, the table's first row will be empty
  With rptWs
    Set tblRange = .Range( _
      .Cells(TABLE_FIRST_ROW, 1), _
      .Cells(rowNum - 1, IIf(fastMode, COL_COUNT - 2, COL_COUNT)) _
      )
  End With
  
  Set rptTable = rptWs.ListObjects.Add( _
    SourceType:=xlSrcRange, _
    Source:=tblRange, _
    XlListObjectHasHeaders:=xlYes)
  
  With rptTable
    .name = TABLE_NAME
    
    With .Range.Font
      .name = "Aptos Narrow"
      .Size = 10
    End With
    
    With .Sort
    
      With .SortFields
        .Clear
        .Add2 key:=rptTable.ListColumns(COL_NUM_NAME).Range
      End With 'SortFields
      
      .Header = xlYes
      .Apply
    End With ' Sort
  End With ' rptTable

  If creatingNewWs Then
    rptWs.Protect
  Else
    ps.Restore rptWs
  End If
  
Exit_Proc:
  OptimizeAppEnvForSpeed False
  mbTitle = REPORT_WORKSHEET_TITLE
  If errorOccurred Then
    ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
    mbPrompt = "Table Worksheet Reference Report generation/update failed"
    mbButtons = vbCritical
  Else
    mbPrompt = "Table Worksheet Reference Report generated/updated in the " & vbCrLf & _
      """" & REPORT_WORKSHEET_NAME & """ Worksheet."
    mbButtons = vbInformation
  End If
  
  If Not silent Then
    MsgBox prompt:=mbPrompt, buttons:=mbButtons, title:=mbTitle
  End If
  
  Exit Sub
Err_Proc:
  errorOccurred = True
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Private Methods
'------------------------------------------------------------------------------'

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

'------------------------------------------------------------------------------'
' Summary: Gets dictionary that represents a list of the code names of the
'   reference report worksheets.
' Remarks: It is usually used to avoid searching the worksheets for defined
'   name or table list object references.
' Return(s): A dictionary that represents a list of the code names of the
'   reference report worksheets.
' Date Created: 2026-07-12
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Private Function GetRefRptCodeNameList() As Dictionary
  Const METHOD_NAME = "GetRefRptCodeNameList"
  
  Dim dict As Dictionary
  Set dict = New Dictionary
  
  dict.Add key:="SheetDefNamesAllRefs", item:=True
  dict.Add key:="SheetDefNamesWsRefsRpt", item:=True
  dict.Add key:="SheetDvAllRpt", item:=True
  dict.Add key:="SheetDvByForumlaRpt", item:=True
  dict.Add key:="SheetTblsAllRefsRpt", item:=True
  dict.Add key:="SheetTblsWsRefsRpt", item:=True

  Set GetRefRptCodeNameList = dict
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Determines if a worksheet is a reference report worksheets using
'   the specified worksheet CodeName property.
' Remarks: It is usually used to avoid searching the worksheets for defined
'   name or table list object references.
' Parameter(s):
'   pCodeName - The Codename property of the worksheet to test.
' Return(s): True if the worksheet specified by its Codename is a reference
'   report worksheet; otherwise False.
' Date Created: 2026-07-12
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Private Function IsWsRefRpt(pCodeName As String) As Boolean
  Const METHOD_NAME = "IsWsRefRpt"
  
  Static dict As Dictionary
  If dict Is Nothing Then
    Set dict = GetRefRptCodeNameList()
  End If

  IsWsRefRpt = dict.Exists(pCodeName)
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

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
