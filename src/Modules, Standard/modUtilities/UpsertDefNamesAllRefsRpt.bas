'------------------------------------------------------------------------------'
' Summary: Updates the defined names all references worksheet.
' Remarks: This sheet includes defined names referenced by:
'   1) Worksheet cells,
'   2, Lambda functions.
'   3) Data validaton sources, and
'   4) VBA code.
' Parameter(s)
'   fast - If True, the source worksheets are not updated; otherwise they
'     are updated with can be very slow. It is an optional parameter with
'     a default value of False.
' Date Created: 2026-06-14
' Date Last Modified: 2026-06-24
'------------------------------------------------------------------------------'
Public Sub UpsertDefNamesAllRefsRpt(Optional fast As Boolean = False)
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "UpsertDefNamesAllRefsRpt"
  
  Const BUTTON_UPDATE_FAST_NAME As String = "btnUpsrtDfNmsAllRfsRptFast"
  Const BUTTON_UPDATE_SLOW_NAME As String = "btnUpsrtDfNmsAllRfsRptSlow"

  Const DEF_NAME_PRINT_TITLES As String = "Print_Titles"
  Const DEST_TABLE_NAME As String = "tbl_Def_Names_All_Refs_Rpt"
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
  Const WS_REF_RPT_TABLE_NAME As String = "tbl_Def_Names_Ws_Refs_Rpt"
  Const TABLE_HEADER_BACKGROUND_COLORINDEX As Long = 8544277
  Const TABLE_HEADER_FOREGROUND_COLORINDEX As Long = 16777215
  Const VAL_RPT_HEADER_CELL_ADDRESS = "Cell_Address"
  Const VAL_RPT_HEADER_FORMULA_SOURCE = "Formula_Source"
  Const VAL_RPT_HEADER_SHEET_NAME = "Sheet_Name"
  Const VAL_RPT_TABLE_NAME As String = "tbl_Data_Validation_Audit_Rpt"
  
  Dim creatingNewDestWs As Boolean
  Dim defName As String
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
  Dim i As Long
  Dim j As Long
  Dim namesCount As Long
  Dim ps As ProtectionState
  Dim refCount As Long
  Dim refCountLimitReached As Boolean
  Dim refCountStr As String
  Dim refCountAllNotZero As Boolean
  Dim refList As String
  Dim rng As Range
  Dim wsRefRptDefNamesTbl As ListObject
  Dim wsRefRptMustBeUpdated As Boolean
  Dim wsRefRptNamesColRng As Range
  Dim wsRefRptRefCntColRng As Range
  Dim wsRefRptReferencesColRng As Range
  Dim wsRefRptRefersToColRng As Range
  Dim wsRefRptVisibleColRng As Range
  Dim valRptTbl As ListObject
  Dim valRptColCellAddressRng As Range
  Dim valRptColFormulaSourceRng As Range
  Dim valRptColSheetNameRng As Range
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
        
  If Not fast Then
    ' This takes a long time to run.
    UpsertDefNamesWsRefRpt fast:=False, silent:=True
  End If
  UpsertDataValAuditRpt silent:=True
  
  ' Optimization: Turn off UI updates
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False
  
  Set destWs = GetWorksheetByCodeName(DEST_WS_CODENAME)
  creatingNewDestWs = (destWs Is Nothing)
  
  If creatingNewDestWs Then
    Set destWs = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count - 1))
    ThisWorkbook.VBProject.VBComponents(destWs.codeName).Name = _
      DEST_WS_CODENAME
      
    With destWs
      .Name = DEST_WS_NAME
      
      With .Cells.Font
        .Name = "Aptos Narrow"
        .Size = 10
      End With
      
      With .Cells(1, 1)
        .Value = DEST_WS_TITLE
        With .Font
          .Name = "Aptos Display"
          .Size = 12
          .Bold = True
        End With
      End With
      
      ' Create the Headers that lie outside the formal table.
      Set rngCollection = New Collection
      
      .Range("B3").Value = "Referencer"
      rngCollection.Add .Range("B3:J3")
      
      .Range("B4").Value = "Worksheets"
      rngCollection.Add .Range("B4:C4")
      
      .Range("D4").Value = "Formulas (Including Lambda)"
      rngCollection.Add .Range("D4:E4")
      
      .Range("F4").Value = "Data Validation Sources"
      rngCollection.Add .Range("F4:G4")
      
      .Range("H4").Value = "VBA Modules"
      rngCollection.Add .Range("H4:I4")
      
      .Range("J4").Value = "All"
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
    Set ps = ProtectionStateStatic.Create(destWs)
    destWs.Unprotect
    
    ' Clear existing ListObject if it exists to avoid conflicts
    On Error Resume Next
    destWs.ListObjects(DEST_TABLE_NAME).Delete
    err.Clear
    On Error GoTo Err_Proc
  End If
  
  With destWs
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_NAME).Value = _
      DEST_COL_NAME_HEADER
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_RWs_COUNT).Value = _
      DEST_COL_RWs_COUNT_HEADER
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_RWs_NAMES).Value = _
      DEST_COL_RWs_NAMES_HEADER
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_RF_COUNT).Value = _
      DEST_COL_RF_COUNT_HEADER
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_RF_NAMES).Value = _
      DEST_COL_RF_NAMES_HEADER
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_RDV_COUNT).Value = _
      DEST_COL_RDV_COUNT_HEADER
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_RDV_NAMES).Value = _
      DEST_COL_RDV_NAMES_HEADER
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_RVBA_COUNT).Value = _
      DEST_COL_RVBA_COUNT_HEADER
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_RVBA_NAMES).Value = _
      DEST_COL_RVBA_NAMES_HEADER
    .Cells(DEST_TABLE_HEADER_ROW, DEST_COL_ANY_REFS).Value = _
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
  Set valRptColCellAddressRng = _
    valRptTbl.ListColumns(VAL_RPT_HEADER_CELL_ADDRESS).DataBodyRange
  Set valRptColFormulaSourceRng = _
    valRptTbl.ListColumns(VAL_RPT_HEADER_FORMULA_SOURCE).DataBodyRange
  Set valRptColSheetNameRng = _
    valRptTbl.ListColumns(VAL_RPT_HEADER_SHEET_NAME).DataBodyRange
    
  namesCount = 0
  For i = 1 To wsRefRptNamesColRng.Count
    defName = wsRefRptNamesColRng(i)
    If wsRefRptVisibleColRng(i).Value Then
      refCountAllNotZero = False
      namesCount = namesCount + 1
      destRow = DEST_TABLE_HEADER_ROW + namesCount
      destWs.Cells(destRow, DEST_COL_NAME).Value = "'" & defName
      
      ' Print titles show up in the Define Names Worksheet References Report
      ' as having no references so we account for them here.
      ' They have format of 'WorksheetName'!Print_Titles
      If Right(defName, Len(DEF_NAME_PRINT_TITLES)) = DEF_NAME_PRINT_TITLES Then
        destWs.Cells(destRow, DEST_COL_RWs_COUNT).Value = 1
        destWs.Cells(destRow, DEST_COL_RWs_NAMES).Value = defName
        refCountAllNotZero = True
        GoTo Skip_For_Print_Title
      End If
      
      destWs.Cells(destRow, DEST_COL_RWs_COUNT).Value = _
        "'" & wsRefRptRefCntColRng(i)
      refCountAllNotZero = refCountAllNotZero Or (wsRefRptRefCntColRng(i) <> 0)
      
      destWs.Cells(destRow, DEST_COL_RWs_NAMES).Value = _
        "'" & wsRefRptReferencesColRng(i)
      
      ' Get the formulas that reference the defined name.
      refCount = 0
      refList = ""
      refCountLimitReached = False
      For j = 1 To wsRefRptNamesColRng.Count
        If j <> i Then
          If ContainsString( _
            searchIn:=wsRefRptRefersToColRng(j), _
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
            
            refList = refList & "{" & wsRefRptRefersToColRng(j) & "}"

          End If
        End If
      Next j
      refCountAllNotZero = refCountAllNotZero Or (refCount > 0)
      refCountStr = _
        "'" & _
        IIf(refCountLimitReached, _
          ">" & REF_COUNT_LIMIT, _
          refCount)
      destWs.Cells(destRow, DEST_COL_RF_COUNT).Value = refCountStr
      destWs.Cells(destRow, DEST_COL_RF_NAMES).Value = "'" & refList
      
      ' Get the data validaton formulas that reference the defined name.
      refCount = 0
      refList = ""
      refCountLimitReached = False
      For j = 1 To valRptColFormulaSourceRng.Count
        If ContainsString( _
          searchIn:=valRptColFormulaSourceRng(j), _
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
          
          refList = _
            refList & _
            "'" & valRptColSheetNameRng(j) & "'!" & _
            valRptColCellAddressRng(j)

        End If
      Next j
      refCountAllNotZero = refCountAllNotZero Or (refCount > 0)
      refCountStr = _
        "'" & _
        IIf(refCountLimitReached, _
          ">" & REF_COUNT_LIMIT, _
          refCount)
      destWs.Cells(destRow, DEST_COL_RDV_COUNT).Value = refCountStr
      destWs.Cells(destRow, DEST_COL_RDV_NAMES).Value = "'" & refList
      
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
      destWs.Cells(destRow, DEST_COL_RVBA_COUNT).Value = refCountStr
      destWs.Cells(destRow, DEST_COL_RVBA_NAMES).Value = "'" & refList
      refCountAllNotZero = refCountAllNotZero Or (refCount > 0)
      
Skip_For_Print_Title:
      destWs.Cells(destRow, DEST_COL_ANY_REFS).Value = refCountAllNotZero
    End If
  Next i

  ' Convert the data range into an Excel Table (ListObject)
  If namesCount = 0 Then
    ' Create the table with a headers row on one blank data row to
    ' avoid an error withthe ListObjects.Add()
    destTblRngStr = "A5:J5"
  Else
    destTblRngStr = "A5:J" & namesCount + 4
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
        .Add2 Key:=destTbl.ListColumns(DEST_COL_NAME).Range
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
      "The Defined Names WORKSHEET Reference Report " & vbCrLf & _
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
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub
