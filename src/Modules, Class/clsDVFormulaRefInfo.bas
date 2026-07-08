Option Explicit

'------------------------------------------------------------------------------'
' Class Module Name: clsDVFormulaRefInfo
' Summary: Contains data validation formula reference information
' Date Created: 2026-07-01
' Date Last Modified: 2026-07-08
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Fields
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Constant Fields
'------------------------------------------------------------------------------'

' The name of this class module.
Private Const MODULE_NAME As String = "clsDVFormulaRefInfo"

'------------------------------------------------------------------------------'
' Variable Fields
'------------------------------------------------------------------------------'

' The validation's Formula1 property.
Private m_F1 As String

' The validation's Formula2 property. A value of "" indicates that its value
' is undefined. It is only defined when the validation Type property is set
' to xlValidateList or xlValidateCustom and the validation Operator is set
' either xlBetween or xlNotBetween.
Private m_F2 As String

' A string the represents the values of both the Validation's Forumula1
' and Formula2 properties. If Formula2 has a value of "" (undefined),
' the string it has the format "Formula1Value"; otherwise, it has the format,
' "1:Formula1Value;2:Formula2Value". For example, if Formula 1 is "x * $A$2",
' and Formula2 is undefined ("") then m_FormulaKey would be "x * $A$2". If
' Formula1 is "12" and Formula2 is "$A$15" then m_FormulaKey would
' "1:12;2:$A$26".
' Remarks: Formula2 is only defined when the validation Type property is
'   set to xlValdateList or xlValidateCustom and the validation Operator
'   is set either xlBetween or xlNotBetween. A value of "" signifies it is
'   undefined.
Private m_FormulaKey As String

' Represents the a list of cells that have use this validation formula. The
' dictionary key is name of the worksheet that contains the cell or cells
' that use this validation formula. The items in the dictionary are
' collections of the cells in the worksheet specified by the dicitionary
' key that use this formula.
' For example if we have the formula, SomeDefinedRange and it is referenced
' by cells a1 and b2 in worksheetOne as well as cells d4 and f6 in
' worksheetTwo we would have:
'
' Set m_RefDict = new Dictionary
'
' Set wsOneCollection =  new Collection
' wsOneCollection.add "a1"
' wsOneCollection.add "b2"
' m_RefDict.add key:="worksheetOne", item:=wsOneCollection
'
' Set wsTwoCollection =  new Collection
' wsOneCollection.add "d4"
' wsOneCollection.add "f6"
' m_RefDict.add key:="worksheetTwo", item:=wsTwoCollection
'
' contains a dictionary of collections of cell addresses keyed by their
' Worksheet names. Only up to the storage count limit references are stored.
Private m_RefDict As Dictionary

'------------------------------------------------------------------------------'
' Pseudo-Constructors
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Summary: Initializes this instance of the clsCellDataValidation with
'   specified property values.
' Parameter(s)
'   pF1 - Represents the cell's Validation Formula1 value.
'   pF2 - Represents the cell's Validation Formula1 value. A value of ""
'     indicates its value is undefined. It is only defined when the validation
'     Type property is set to xlValidateList or xlValidateCustom and the
'     validation Operator is set either xlBetween or xlNotBetween.
'   pRefDict - See m_RefDictIt above for a description. It is an optional
'   parameter with a default value of nothing.
' Date Created: 2026-06-29
' Date Last Modified: 2026-07-07
'---------------------------------------------------s---------------------------'
Friend Sub Init( _
  ByVal pF1 As String, _
  ByVal pF2 As String, _
  Optional ByRef pRefDict As Dictionary = Nothing)
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "Init"
  
  m_F1 = pF1
  m_F2 = pF2
  m_FormulaKey = CombineF1F2()
  If pRefDict Is Nothing Then
    Set m_RefDict = New Dictionary
  Else
    Set m_RefDict = pRefDict
  End If

Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Initializes this instance of the clsCellDataValidation with the
'   validation information from the specified cell.
' Parameter(s)
'   pTargetCell - The cell whose validation information is used.
'   pRefDict - Represents the list of cells that have use this validation
'     formula. It contains a dictionary of collections of cell addresses
'     keyed by their Worksheet names.It is an optional parameter with a default value
'     of nothing.
' Date Created: 2026-07-07
' Date Last Modified: 2026-07-07
'---------------------------------------------------s---------------------------'
Friend Sub InitWithCell( _
  ByRef pCell As Range, _
  Optional ByRef pRefDict As Dictionary = Nothing _
  )
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "InitWithCell"
  
  Dim f1 As String
  Dim f1Type As ValidationSrcType
  Dim f2 As String
  Dim f2Type As ValidationSrcType
  Dim op As XlFormatConditionOperator
  
  GetDvFormulasAndTypes _
    targetCell:=pCell, _
    op:=op, _
    f1:=f1, _
    f1Type:=f1Type, _
    f2:=f2, _
    f2Type:=f2Type
    
  Init _
    pF1:=f1, _
    pF2:=f2, _
    pRefDict:=pRefDict

Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Properties
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Summary: Gets the Formula1 property's value.
' Value: The validation's Formula1 property. It is a readonly property.
' Date Created: 2026-07-01
' Date Last Modified: 2026-07-07
'------------------------------------------------------------------------------'
Friend Property Get Formula1() As String
  Formula1 = m_F1
End Property


'------------------------------------------------------------------------------'
' Summary: Gets the Formula2 property's value.
' Value: The validation's Formula1 property. A value of "" indicates it is not
'   defined. it is a readonly property.
' Date Created: 2026-07-01
' Date Last Modified: 2026-07-07
'------------------------------------------------------------------------------'
Friend Property Get formula2() As String
  formula2 = m_F2
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the FormulaKey property's value.
' Value:  See m_FormulaKey above for a description. It is a readonly property
'   that is initialized by the pseudo-constructor using the values of Formula1
'   and Formula2.
' Date Created: 2026-07-07
' Date Last Modified: 2026-07-07
'------------------------------------------------------------------------------'
Friend Property Get formulaKey() As String
  formulaKey = m_FormulaKey
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the RefDict property's value.
' Value:  See m_RefDict for a description.
' Date Created: 2026-07-01
' Date Last Modified: 2026-07-07
'------------------------------------------------------------------------------'
Friend Property Get RefDict() As Dictionary
  Set RefDict = m_RefDict
End Property

'------------------------------------------------------------------------------'
' Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Friend Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Summary: Add a reference to the reference dictionary.
' Parameter(s)
'   wsName - The name of the Worksheet that contains the cell that uses this
'    validation formula.
'   cellAddress- The address of the cell that uses this validation formula.
' Date Created: 2026-07-02
' Date Last Modified: 2026-07-07
'------------------------------------------------------------------------------'
Friend Sub AddReference(wsName As String, cellAddress As String)
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "AddReference"

  Dim cellAddressCollection As Collection

  If Not m_RefDict.Exists(wsName) Then
    Set cellAddressCollection = New Collection
    cellAddressCollection.Add item:=cellAddress
    Set m_RefDict = New Dictionary
    m_RefDict.Add key:=wsName, item:=cellAddressCollection
  Else
    Set cellAddressCollection = m_RefDict(wsName)
    cellAddressCollection.Add item:=cellAddress
  End If

Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

'------------------------------------------------------------------------------'
' Summary: Gets the count of cells that use this validation formula.
' Return(s): Get the count of cells that reference the formula represented by
'   this instance of the clsDVFormulaRefInfo class.
' Date Created: 2026-07-02
' Date Last Modified: 2026-07-08
'------------------------------------------------------------------------------'
Friend Function GetRefCount() As Long
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "GetRefCount"
  
  Dim cellAddressCollection As Collection
  Dim item As Variant
  
  GetRefCount = 0
  For Each item In m_RefDict.Items
    Set cellAddressCollection = item
    GetRefCount = GetRefCount + cellAddressCollection.count
  Next item
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Gets a string that represents the cells that reference the formula
'   represented by this instance of the clsDVFormulaRefInfo class.
' Remarks: It has the format:
'   "'Ws_Name1'!CellAddrA,CellAddrB;'Ws_Name2'!CellAddrC;WsName3'!CellAddrD".
'   The count of cell references exceeds the cell count limit then "..."
'   is appended to the string.
' Parameter(s):
'   refLimitExceeded - Indicates that the count of cell refences exceeds the
'     reference count limit and consequently, the number of references shown
'     in the returned string is equal to the reference count limit and the
'     the complete list of cell references in not contained in the returned
'     string.
'   refLimit - Represents a limit to the count of cell references represented
'     by the returned string. If value is 0 then, an unlimited count of
'     references are used. It is an optional parameter with a default value
'      of 10.
' Returns: Gets a string that represents the cells that reference the formula
'   represented by this instance of the clsDVFormulaRefInfo class. It is
'   subject to the limits explained above.
' Date Created: 2026-07-02
' Date Last Modified: 2026-07-07
'------------------------------------------------------------------------------'
Friend Function GetRefString( _
  refLimitExceeded As Boolean, _
  Optional refLimit As Long = 10) _
  As String
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "GetRefString"
  
  Dim cellAddress As Variant
  Dim cellAddressCollection As Collection
  Dim firstAddress As Boolean
  Dim firstWorksheet As Boolean
  Dim refCount As Long
  Dim refString As String
  Dim worksheetKey As Variant
  
  refCount = 0
  refLimitExceeded = False
  firstWorksheet = True
  
  For Each worksheetKey In m_RefDict
    firstAddress = True
    Set cellAddressCollection = m_RefDict(worksheetKey)
    For Each cellAddress In cellAddressCollection
      If refLimit = 0 Or refCount < refLimit Then
        ' If we are on the first address, add the worksheet name.
        If firstAddress Then
          ' If we are not on the first worksheet, 1) add
          ' the worksheet separator, ";"
          If Not firstWorksheet Then
            refString = refString & ";"
            firstWorksheet = False
          End If
          ' Then 2) add the worksheet name.
          refString = refString & "'" & worksheetKey & "'!"
          firstAddress = False
        Else ' we are not on the first address, so add the address separator, ","
          refString = refString & ","
        End If
        ' Then add the address.
        refString = refString & cellAddress
        refCount = refCount + 1
      Else ' We have exceeded the reference count limit.
        refLimitExceeded = True
        refString = refString & ",..."
        GoTo Next_Worksheet_Key_Iter:
      End If
    Next cellAddress
    
Next_Worksheet_Key_Iter:
  Next worksheetKey
  GetRefString = refString
  
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
' Summary: Creates a string the represents the values of both the
'   Validation's Forumula1 and Formula2 properties. If Formula2 has a value
'   of "" (undefined_), it has the format "Formula1Value"; otherwise, it has
'   the format "1:Formula1Value;2:Formula2Value".
' Remarks: Formula2 is only defined when the validation Type property is
'   set to xlValidateList or xlValidateCustom and the validation Operator
'   is set either xlBetween or xlNotBetween. A value of "" signifies it is
'   undefined.
' Date Created: 2026-07-07
' Date Last Modified: 2026-07-07
'------------------------------------------------------------------------------'
Private Function CombineF1F2() As String
  
  On Error GoTo Err_Proc
  
  Const METHOD_NAME As String = "CombineF1F2"
  Dim result As String
  
  If m_F2 = "" Then
    result = m_F1
  Else
    result = "1:" & m_F1 & ";2:" & m_F2
  End If
  
  CombineF1F2 = result
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function
