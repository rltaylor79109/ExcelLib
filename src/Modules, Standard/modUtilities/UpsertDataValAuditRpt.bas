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
' Date Last Modified: 2026-06-24
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

  Dim creatingNewWs As Boolean
  Dim errorOccurred As Boolean
  Dim ps As ProtectionState
  Dim rowCount As Long
  Dim rptTable As ListObject
  Dim rptWs As Worksheet
  Dim targetCell As Range
  Dim tblRange As Range
  Dim updateButton As Button
  Dim validationRange As Range
  Dim ws As Worksheet
  Dim wb As Workbook

  errorOccurred = False
  
  ' Optimization: Turn off UI updates
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False
  
  Set wb = ThisWorkbook
  
  ' Create or clear the audit sheet
  Set rptWs = _
    GetWorksheetByCodeName(REPORT_WORKSHEET_CODENAME)
  creatingNewWs = (rptWs Is Nothing)
  
  If creatingNewWs Then
    Set rptWs = wb.Worksheets.Add(After:=Sheets(Sheets.Count - 1))
    ThisWorkbook.VBProject.VBComponents(rptWs.codeName).Name = _
      REPORT_WORKSHEET_CODENAME
    
    With rptWs
      .Name = REPORT_WORKSHEET_NAME
      
      With .Cells.Font
        .Name = "Aptos Narrow"
        .Size = 10
      End With
      
      With .Cells(1, 1)
        .Value = REPORT_WORKSHEET_TITLE
        With .Font
          .Name = "Aptos Display"
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
      .Name = BUTTON_NAME
      .Placement = xlFreeFloating
    End With
  Else
    Set ps = ProtectionStateStatic.Create(rptWs)
    rptWs.Unprotect
    ' Clear existing ListObject if it exists to avoid conflicts
    On Error Resume Next
    rptWs.ListObjects(TABLE_NAME).Delete
    err.Clear
    On Error GoTo Err_Proc
  End If
    
  ' Headers\
  rptWs.Cells(TABLE_FIRST_ROW, COL_INDEX_FORMULA_SOURCE).Value = _
    COL_NAME_FORMULA_SOURCE
  rptWs.Cells(TABLE_FIRST_ROW, COL_INDEX_SHEET_INDEX).Value = _
    COL_NAME_SHEET_NAME
  rptWs.Cells(TABLE_FIRST_ROW, COL_INDEX_CELL_ADDRESS).Value = _
    COL_NAME_CELL_ADDRESS
  rptWs.Cells(TABLE_FIRST_ROW, COL_INDEX_VALIDATION_TYPE).Value = _
    COL_NAME_VALIDATION_TYPE
  rptWs.Range( _
    rptWs.Cells(TABLE_FIRST_ROW, 1), _
    rptWs.Cells(TABLE_FIRST_ROW, COL_COUNT) _
    ).Font.Bold = True
  rowCount = TABLE_FIRST_ROW + 1

  ' Loop through all sheets to find Validation
  For Each ws In ThisWorkbook.Worksheets
    If ws.codeName <> REPORT_WORKSHEET_CODENAME Then
      Set validationRange = Nothing
      ' SpecialCells will error if NO cells have validation, so we keep this one
      On Error Resume Next
      Set validationRange = ws.Cells.SpecialCells(xlCellTypeAllValidation)
      err.Clear
      On Error GoTo Err_Proc
      
      If Not validationRange Is Nothing Then
        For Each targetCell In validationRange.Cells
          Dim vType As Long
          vType = -1
          On Error Resume Next
          vType = targetCell.Validation.Type
          err.Clear
          On Error GoTo Err_Proc
          
          ' Type 3 is xlValidateList
          If vType = 3 Then
            rptWs.Cells(rowCount, 1).Value = "'" & targetCell.Validation.Formula1
            rptWs.Cells(rowCount, 2).Value = "'" & ws.Name
            rptWs.Cells(rowCount, 3).Value = "'" & targetCell.Address
            rptWs.Cells(rowCount, 4).Value = "'" & "List"
            ' Prepend ' to ensure formulas are treated as text
            rowCount = rowCount + 1
          End If
        Next targetCell
      End If
    End If
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
      .Name = TABLE_NAME
      
      With .Range.Font
        .Name = "Aptos Narrow"
        .Size = 10
      End With ' .Range.Font
  
    ' Sorting Logic using the newly created Table
      With .Sort
      
        With .SortFields
          .Clear
          .Add2 Key:=rptTable.ListColumns(COL_NAME_FORMULA_SOURCE).Range
          .Add2 Key:=rptTable.ListColumns(COL_NAME_SHEET_NAME).Range
          .Add2 Key:=rptTable.ListColumns(COL_NAME_CELL_ADDRESS).Range
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
    ps.RestoreProtection rptWs
  End If
  
Exit_Proc:
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
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
