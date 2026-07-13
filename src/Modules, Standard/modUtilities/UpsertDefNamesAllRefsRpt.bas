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
' Date Last Modified: 2026-07-13
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
  
  ' Optimization: Turn off UI updates
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False
  
  Set destWs = GetWorksheetByCodeName(DEST_WS_CODENAME)
  creatingNewDestWs = (destWs Is Nothing)
  
  If creatingNewDestWs Then
    Set destWs = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.count - 1))
    ThisWorkbook.VBProject.VBComponents(destWs.codeName).Name = _
      DEST_WS_CODENAME
      
    With destWs
      .Name = DEST_WS_NAME
      
      With .cells.Font
        .Name = "Aptos Narrow"
        .Size = 10
      End With
      
      With .cells(1, 1)
        .value = DEST_WS_TITLE
        With .Font
          .Name = "Aptos Display"
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
      .Name = BUTTON_UPDATE_SLOW_NAME
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
      .Name = BUTTON_UPDATE_FAST_NAME
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
    .cells(DEST_TABLE_HEADER_ROW, DEST_COL_NAME).value = _
      DEST_COL_NAME_HEADER
    .cells(DEST_TABLE_HEADER_ROW, DEST_COL_RWs_COUNT).value = _
      DEST_COL_RWs_COUNT_HEADER
    .cells(DEST_TABLE_HEADER_ROW, DEST_COL_RWs_NAMES).value = _
      DEST_COL_RWs_NAMES_HEADER
    .cells(DEST_TABLE_HEADER_ROW, DEST_COL_RF_COUNT).value = _
      DEST_COL_RF_COUNT_HEADER
    .cells(DEST_TABLE_HEADER_ROW, DEST_COL_RF_NAMES).value = _
      DEST_COL_RF_NAMES_HEADER
    .cells(DEST_TABLE_HEADER_ROW, DEST_COL_RDV_COUNT).value = _
      DEST_COL_RDV_COUNT_HEADER
    .cells(DEST_TABLE_HEADER_ROW, DEST_COL_RDV_NAMES).value = _
      DEST_COL_RDV_NAMES_HEADER
    .cells(DEST_TABLE_HEADER_ROW, DEST_COL_RVBA_COUNT).value = _
      DEST_COL_RVBA_COUNT_HEADER
    .cells(DEST_TABLE_HEADER_ROW, DEST_COL_RVBA_NAMES).value = _
      DEST_COL_RVBA_NAMES_HEADER
    .cells(DEST_TABLE_HEADER_ROW, DEST_COL_ANY_REFS).value = _
      DEST_COL_ANY_REFS_HEADER
    .Range( _
      .cells(DEST_TABLE_HEADER_ROW, 1), _
      .cells(DEST_TABLE_HEADER_ROW, DEST_TABLE_COL_COUNT) _
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
      destWs.cells(destRow, DEST_COL_NAME).value = "'" & defName
      
      ' Print titles show up in the Define Names Worksheet References Report
      ' as having no references so we account for them here.
      ' They have format of 'WorksheetName'!Print_Titles
      If Right(defName, Len(DEF_NAME_PRINT_TITLES)) = DEF_NAME_PRINT_TITLES Then
        destWs.cells(destRow, DEST_COL_RWs_COUNT).value = 1
        destWs.cells(destRow, DEST_COL_RWs_NAMES).value = defName
        refCountAllNotZero = True
        GoTo Skip_For_Print_Title
      End If
      
      destWs.cells(destRow, DEST_COL_RWs_COUNT).value = _
        "'" & wsRefRptRefCntColRng(defNameDestIdx)
      refCountAllNotZero = refCountAllNotZero Or (wsRefRptRefCntColRng(defNameDestIdx) <> 0)
      
      destWs.cells(destRow, DEST_COL_RWs_NAMES).value = _
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
      destWs.cells(destRow, DEST_COL_RF_COUNT).value = refCountStr
      destWs.cells(destRow, DEST_COL_RF_NAMES).value = "'" & refList
      
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
      destWs.cells(destRow, DEST_COL_RDV_COUNT).value = refCount
      Dim refListStr As String
      If refCount > 0 Then
        refListStr = CombineAndTruncateRefs( _
          refStrings:=refListCollection, _
          refCountLimit:=REF_COUNT_LIMIT)
      Else
        refListStr = ""
      End If
      
      destWs.cells(destRow, DEST_COL_RDV_NAMES).value = "'" & refListStr
      
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
          
          refList = refList & vbComp.Name
          
        End If
      Next vbComp

      refCountStr = _
        "'" & _
        IIf(refCountLimitReached, _
          ">" & REF_COUNT_LIMIT, _
          refCount)
      destWs.cells(destRow, DEST_COL_RVBA_COUNT).value = refCountStr
      destWs.cells(destRow, DEST_COL_RVBA_NAMES).value = "'" & refList
      refCountAllNotZero = refCountAllNotZero Or (refCount > 0)
      
Skip_For_Print_Title:
      destWs.cells(destRow, DEST_COL_ANY_REFS).value = refCountAllNotZero
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
    .Name = DEST_TABLE_NAME
    
    With .Range.Font
      .Name = "Aptos Narrow"
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
    ps.RestoreProtection destWs
  End If
  
Exit_Proc:
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
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
