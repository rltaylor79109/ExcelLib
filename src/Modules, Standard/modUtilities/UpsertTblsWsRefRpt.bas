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
' Date Last Modified: 2026-07-13
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
  
  ' Optimization: Turn off UI updates
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False
  
  Set wb = ThisWorkbook
  
  ' Create or clear the report worksheet
  Set rptWs = GetWorksheetByCodeName(REPORT_WORKSHEET_CODENAME)
  creatingNewWs = (rptWs Is Nothing)
  
  If creatingNewWs Then
    Set rptWs = wb.Worksheets.Add(After:=Sheets(Sheets.count - 1))
    rptWs.Name = REPORT_WORKSHEET_NAME
    ' This requires "Trust access to the VBA project object model" to be enabled and
    ' also for the workbook file to unblocked.
    ThisWorkbook.VBProject.VBComponents(rptWs.codeName).Name = REPORT_WORKSHEET_CODENAME

    With rptWs.cells.Font
        .Name = "Aptos Narrow"
        .Size = 10
    End With
      
    With rptWs.cells(1, 1)
      .value = REPORT_WORKSHEET_TITLE
      With .Font
        .Name = "Aptos Display"
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
      .Name = SLOW_BUTTON_NAME
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
      .Name = FAST_BUTTON_NAME
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
    .cells(TABLE_FIRST_ROW, COL_NUM_NAME) = COL_HEADER_NAME
    .cells(TABLE_FIRST_ROW, COL_NUM_ADDRESS) = COL_HEADER_ADDRESS
    .cells(TABLE_FIRST_ROW, COL_NUM_COMMENT) = COL_HEADER_COMMENT
    .cells(TABLE_FIRST_ROW, COL_NUM_REF_COUNT) = _
      IIf(fastMode, "", COL_HEADER_REF_COUNT)
    .cells(TABLE_FIRST_ROW, COL_NUM_REFS) = IIf(fastMode, "", COL_HEADER_REFS)
    .Range( _
      .cells(TABLE_FIRST_ROW, 1), _
      .cells(TABLE_FIRST_ROW, COL_COUNT) _
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
      rptWs.cells(rowNum, COL_NUM_NAME).value = "'" & targetTbl.Name
           
      ' Table Range Address
      rptWs.cells(rowNum, COL_NUM_ADDRESS).value = _
        "'" & "'" & wsForTblSearch.Name & "'!" & targetTbl.Range.Address
        
      ' Comment
      rptWs.cells(rowNum, COL_NUM_COMMENT) = "'" & targetTbl.Comment
      
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
          
          inStrResult = InStr(1, targetCell.formula, targetTbl.Name, vbTextCompare)
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
            refList = refList & "'" & wsForRefSearch.Name & "'!"
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
      rptWs.cells(rowNum, COL_NUM_REF_COUNT).value = "'" & refCountStr
      rptWs.cells(rowNum, COL_NUM_REFS).value = "'" & refList
      
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
      .cells(TABLE_FIRST_ROW, 1), _
      .cells(rowNum - 1, IIf(fastMode, COL_COUNT - 2, COL_COUNT)) _
      )
  End With
  
  Set rptTable = rptWs.ListObjects.Add( _
    SourceType:=xlSrcRange, _
    Source:=tblRange, _
    XlListObjectHasHeaders:=xlYes)
  
  With rptTable
    .Name = TABLE_NAME
    
    With .Range.Font
      .Name = "Aptos Narrow"
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
    ps.RestoreProtection rptWs
  End If
  
Exit_Proc:
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
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
