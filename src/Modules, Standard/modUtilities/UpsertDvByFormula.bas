'------------------------------------------------------------------------------'
' Summary: Creates or updates a worksheet that shows data validation
'   formulas.
' Parameter(s)
'   silent - It True, the method does not display any messages to the user
'     except for errors; otherwise a report updated message is displayed to
'     the user when complete. It is an optional parameter with a default value
'     of False.'
' Date Created: 2026-07-07
' Date Last Modified: 2026-07-12
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

      With .cells.Font
        .Name = "Aptos Narrow"
        .Size = 10
      End With

      With .cells(1, 1)
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
      .OnAction = "UpsertDvByFormulaRpt"
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
  rptWs.cells(TBL_FIRST_ROW, COL_IDX_F1).value = _
    COL_HDR_F1 ' Col 1
  rptWs.cells(TBL_FIRST_ROW, COL_IDX_F2).value = _
    COL_HDR_F2 ' Col 2
  rptWs.cells(TBL_FIRST_ROW, COL_IDX_REF_COUNT).value = _
    COL_HDR_REF_COUNT ' Col 3
  rptWs.cells(TBL_FIRST_ROW, COL_IDX_REFS).value = _
    COL_HDR_REFS ' Col 4
    
  rptWs.Range( _
    rptWs.cells(TBL_FIRST_ROW, 1), _
    rptWs.cells(TBL_FIRST_ROW, COL_COUNT) _
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
    Set valRng = ws.cells.SpecialCells(xlCellTypeAllValidation)
    err.Clear
    On Error GoTo Err_Proc
    
    ' If no cells in this worksheet have defined data validation, then
    ' go to next worksheet iteration.
    If valRng Is Nothing Then GoTo Continue_ws
    
    For Each targetCell In valRng.cells
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
        wsName:=ws.Name, _
        cellAddress:=targetCell.Address
    Next targetCell
Continue_ws:
  Next ws
      
  For Each varFormulaKey In formulaRefInfoDict.Keys
    Set formulaRefInfo = formulaRefInfoDict(varFormulaKey)
    rptWs.cells(rowCount, COL_IDX_F1).value = _
      "'" & formulaRefInfo.Formula1 ' col 1
    rptWs.cells(rowCount, COL_IDX_F2).value = _
      "'" & formulaRefInfo.formula2 ' col 2
    rptWs.cells(rowCount, COL_IDX_REF_COUNT).value = _
      formulaRefInfo.GetRefCount ' col 3
    rptWs.cells(rowCount, COL_IDX_REFS).value = _
      "'" & formulaRefInfo.GetRefString( _
        refLimitExceeded:=refCountLimitExceeded, _
        refLimit:=REF_COUNT_LIMIT) ' col 4
    rowCount = rowCount + 1
  Next varFormulaKey

  ' Convert Data Range into an Excel Table (ListObject)
  If rowCount > TBL_FIRST_ROW + 2 Then
    Set tblRange = rptWs.Range( _
      rptWs.cells(TBL_FIRST_ROW, 1), _
      rptWs.cells(rowCount - 1, COL_COUNT))
  ' If no data validation was found, create table with header
  ' and one blank row.
  Else
    Set tblRange = rptWs.Range( _
      rptWs.cells(TBL_FIRST_ROW, 1), _
      rptWs.cells(TBL_FIRST_ROW + 1, COL_COUNT))
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
    ps.RestoreProtection rptWs
  End If

Exit_Proc:
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
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
