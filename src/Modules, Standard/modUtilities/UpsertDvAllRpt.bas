'------------------------------------------------------------------------------'
' Summary: Creates or updates a worksheet that shows data validation
'   properties for all cells in the workbook that have data validation defined.
' Parameter(s)
'   silent - It True, the method does not display any messages to the user
'     except for errors; otherwise a report updated message is displayed to
'     the user when complete. It is an optional parameter with a default value
'     of False.'
' Date Created: 2026-07-07
' Date Last Modified: 2026-07-07
'------------------------------------------------------------------------------'
Public Sub UpsertDvAllRpt(Optional silent As Boolean = False)
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "UpsertDvAllRpt"

  Const BUTTON_NAME As String = "btnUpsertDvAllRptt"
  
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
  Const TBL_NAME As String = "tbl_Dv_All_Rpt"

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

  ' Optimization: Turn off UI updates
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False

  Set wb = ThisWorkbook

  ' Create or clear the audit sheet
  Set rptWs = _
    GetWorksheetByCodeName(RPT_WS_CODENAME)
  creatingNewWs = (rptWs Is Nothing)

  If creatingNewWs Then
    Set rptWs = wb.Worksheets.Add(After:=Sheets(Sheets.count - 1))
    ThisWorkbook.VBProject.VBComponents(rptWs.codeName).Name = _
      RPT_WS_CODENAME

    With rptWs
      .Name = RPT_WS_NAME

      With .Cells.Font
        .Name = "Aptos Narrow"
        .Size = 10
      End With

      With .Cells(1, 1)
        .value = RPT_WS_TITLE
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
      .OnAction = "RunUpsertDvAllRpt"
      .Caption = "Update"
      .Name = BUTTON_NAME
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
    If ws.codeName = RPT_WS_CODENAME Then GoTo NEXT_WS_ITER
    
    Set valRng = Nothing
    ' An attempt reference to Range.SpecialCells(xlCellTypeAllValidation)
    ' will raise an error if no cells in the range have validation.
    On Error Resume Next
    Set valRng = ws.Cells.SpecialCells(xlCellTypeAllValidation)
    err.Clear
    On Error GoTo Err_Proc
    
    ' If no cells in this worksheet have defined data validation, then
    ' go to next worksheet iteration.
    If valRng Is Nothing Then GoTo NEXT_WS_ITER
    
    For Each targetCell In valRng.Cells
      ' Gemini suggested error checking here but the Set valRng statement
      ' above should insure each cell in the range has validation defined.

      rptWs.Cells(rowCount, COL_IDX_SHEET).value = _
        "'" & targetCell.Parent.Name ' Col 1
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
    
NEXT_WS_ITER:
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
    .Name = TBL_NAME

    With .Range.Font
      .Name = "Aptos Narrow"
      .Size = 10
    End With ' .Range.Font

  ' Sorting Logic using the newly created Table
    With .Sort

      With .SortFields
        .Clear
        .Add2 Key:=rptTable.ListColumns(COL_IDX_SHEET).Range
        .Add2 Key:=rptTable.ListColumns(COL_IDX_CELL).Range
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
    ps.RestoreProtection rptWs
  End If

Exit_Proc:
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
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