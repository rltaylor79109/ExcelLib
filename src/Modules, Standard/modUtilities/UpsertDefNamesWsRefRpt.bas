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
' Date Last Modified: 2026-07-13
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
  Dim defName As Name

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
  Dim testName As Name

  Dim updateButtonFast As Button
  Dim updateButtonSlow As Button
  Dim wb As Workbook
  Dim ws As Worksheet
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
      .OnAction = "RunUpsrtDfNmsWsRptSlowVrbs"
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
      .OnAction = "RunUpsrtDfNmsWsRptFastVrbs"
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
    .cells(TABLE_FIRST_ROW, COL_NUM_REFERS_TO) = COL_HEADER_REFERS_TO
    .cells(TABLE_FIRST_ROW, COL_NUM_SCOPE) = COL_HEADER_SCOPE
    .cells(TABLE_FIRST_ROW, COL_NUM_VALUE) = COL_HEADER_VALUE
    .cells(TABLE_FIRST_ROW, COL_NUM_VISIBLE) = COL_HEADER_VISIBLE
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
    rptWs.cells(rowNum, COL_NUM_NAME).value = "'" & defName.Name
         
    ' Use RefersTo, the defined name definition.
    rptWs.cells(rowNum, COL_NUM_REFERS_TO).value = "'" & defName.RefersTo

    ' Scope
    If defName.Parent Is wb Then
      rptWs.cells(rowNum, COL_NUM_SCOPE).value = "'Workbook"
    Else
      rptWs.cells(rowNum, COL_NUM_SCOPE).value = "'" & defName.Parent.Name
    End If
      
    ' Value
    nameValue = Application.Evaluate(defName.RefersTo)
    If IsError(nameValue) Then
        rptWs.cells(rowNum, COL_NUM_VALUE).value = "'Error or Invalid"
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
        rptWs.cells(rowNum, COL_NUM_VALUE).value = "'" & valueString
      Else
        rptWs.cells(rowNum, COL_NUM_VALUE).value = "'Array (Non-standard)"
      End If
    Else
      ' Handle single values or formulas
      rptWs.cells(rowNum, COL_NUM_VALUE).value = "'" & nameValue
    End If
    
    ' Visible
    rptWs.cells(rowNum, COL_NUM_VISIBLE) = defName.Visible
    
    ' Comment
    rptWs.cells(rowNum, COL_NUM_COMMENT) = "'" & defName.Comment
    
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
        
        inStrResult = InStr(1, aCell.formula, defName.Name, vbTextCompare)
        
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
          refList = refList & "'" & ws.Name & "'!"
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
    rptWs.cells(rowNum, COL_NUM_REF_COUNT).value = "'" & refCountStr
    rptWs.cells(rowNum, COL_NUM_REFS).value = "'" & refList
    
Skip_Reference_Search:
    rowNum = rowNum + 1

Continue_WbDefName:
  Next defName
  
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
    ps.RestoreProtection rptWs
  End If
  
Exit_Proc:
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
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
