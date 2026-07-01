Option Explicit

'------------------------------------------------------------------------------'
' Class Module Name: clsDVFormulaRefInfo
' Summary: Contains data validation formula reference information
' Date Created: 2026-07-01
' Date Last Modified: 2026-07-01
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Fields
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Constant Fields
'------------------------------------------------------------------------------'

' The name of this class module.
Private Const MODULE_NAME As String = "clsDVFormulaRefInfo"

' The reference count storage limit. If the data validation has more
' references than this limit, we only store up to this limit and flag the
' limit has been exceeded.
Private Const REF_COUNT_LIMIT As Long = 10

'------------------------------------------------------------------------------'
' Variable Fields
'------------------------------------------------------------------------------'

' The validation's Formula1 property.
Private m_Formula1 As String

' The validation's Formula2 property. A value of "" indicates that its value
' is undefined. It is only defined when the validation Type property is set
' to xlValidateList or xlValidateCustom and the validation Operator is set
' either xlBetween or xlNotBetween.
Private m_Formula2 As String

' The count of cells that use this validaiton formula. If
' m_RefCountExceedsLimit is true this count is invalid and the formula is
' used by more than REF_COUNT_LIMIT cells.
Private m_RefCount As Long

' Indicates whether the count of cells that use this validation formula
' exceeds the reference count storage limit.
Private m_RefCountExceedsLimit As Boolean

' Represents the list of cells that have use this validation formula. It has
' format of "'WsName1'!Addr1,Addr2,...;'WsName2'!Addr1,Addr2,...;..."
' where WsName is the Worksheet Name and Addr is the cell address.
' Only the first REF_COUNT_LIMIT references are stored.
Private m_RefList As String

'------------------------------------------------------------------------------'
' Pseudo-Constructors
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Summary: Instantiates and initializes an instance of the
'    clsCellDataValidation with specified property values
' Parameter(s)
'   formula1 - Represents the cell's Validation Formula1 value.
'   formula2 - Represents the cell's Validation Formula1 value. A value of ""
'     indicates its value is undefined. It is only defined when the validation
'     Type property is set to xlValidateList or xlValidateCustom and the
'     validation Operator is set either xlBetween or xlNotBetween.
'   refCount - Represent the count of cells that use this validaiton formula.
'     If m_RefCountExceedsLimit is True this count is invalid and the formula
'     is used by more than REF_COUNT_LIMIT cells. It an optional parameter
'     with a default value of 0.
'   refCountExceedsLimit - Indicates whether the the count of cells that use
'     this validation formula exceeds the reference count storage limit. It an
'     optional parameter with a default value of False.
'   refList - Represents the list of cells that have use this validation
'     formula. It has the format
'     "'WsName1'!Addr1,Addr2,...;'WsName2'!Addr1,Addr2,...;..." where WsName
'     is the Worksheet Name and AddrN are the cell addresses. Only the first
'     REF_COUNT_LIMIT references are stored.
' Date Created: 2026-06-29
' Date Last Modified: 2026-07-01
'---------------------------------------------------s---------------------------'
Friend Sub Init( _
  ByVal formula1 As String, _
  ByVal formula2 As String, _
  Optional ByVal refCount As Long = 0, _
  Optional ByVal refCountExceedsLimit As Boolean = False, _
  Optional ByVal refList As String = "")
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "Init"
  
  m_Formula1 = formula1
  m_Formula2 = formula2
  m_RefCount = refCount
  m_RefCountExceedsLimit = refCountExceedsLimit
  m_RefList = refList

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
' Value: The validation's Formula1 property.
' Date Created: 2026-07-01
' Date Last Modified: 2026-07-01
'------------------------------------------------------------------------------'
Friend Property Get formula1() As Boolean
  formula1 = m_Formula1
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the Formula2 property's value.
' Value: The validation's Formula1 property. A value of "" indicates it is not
'   defined.
' Remarks: Formula2 is only defined when the validation Type property is
'   set to xlValidateList or xlValidateCustom and the validation Operator
'   is set either xlBetween or xlNotBetween.
' Date Created: 2026-07-01
' Date Last Modified: 2026-07-01
'------------------------------------------------------------------------------'
Friend Property Get formula2() As Boolean
  formula2 = m_Formula2
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the FormulasBoth property's value.
' Value: A strings the represents the values of both the Forumula1 and
'   Formula2 properties. If Formula2 has a value of "" (undefined_), it has
'   the format "Formula1Value"; otherwise, it has the format
'   "1:Formula1Value;2:Formula2Value".
' Remarks: Formula2 is only defined when the validation Type property is
'   set to xlValidateList or xlValidateCustom and the validation Operator
'   is set either xlBetween or xlNotBetween. A value of "" signifies it is
'   undefined.
' Date Created: 2026-07-01
' Date Last Modified: 2026-07-01
'------------------------------------------------------------------------------'
Friend Property Get FormulasBoth() As String
  Dim value As String
  
  If m_Formula2 = "" Then
    value = m_Formula1
  Else
    value = "1:" & m_Formula1 & ";2:" & m_Formula2
  End If
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the RefCount property's value.
' Value: The count of cells that use this validaiton formula. If
' m_RefCountExceedsLimit is true this count is invalid and the formula is
' used by more than REF_COUNT_LIMIT cells.
' Date Created: 2026-07-01
' Date Last Modified: 2026-07-01
'------------------------------------------------------------------------------'
Friend Property Get refCount() As Long
  refCount = m_RefCount
End Property
Public Property Let refCount(ByVal value As Long)
  m_RefCount = value
End Property


'------------------------------------------------------------------------------'
' Summary: Gets the RefCountExceedsLimit property's value.
' Value: Indicates whether the the count of cells that use
'     this validation formula exceeds the reference count storage limit.
' Date Created: 2026-07-01
' Date Last Modified: 2026-07-01
'------------------------------------------------------------------------------'
Friend Property Get refCountExceedsLimit() As Boolean
  refCountExceedsLimit = m_RefCountExceedsLimit
End Property
Public Property Let refCountExceedsLimit(ByVal value As Boolean)
  m_RefCountExceedsLimit = value
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the RefCountLimit property's value.
' Value: The reference count storage limit. If the data validation has more
'   references than this limit, we only store up to this limit and flag that
'   the limit has been exceeded.
' Date Created: 2026-07-01
' Date Last Modified: 2026-07-01
'------------------------------------------------------------------------------'
Friend Property Get RefCountLimit() As Long
  RefCountLimit = REF_COUNT_LIMIT
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the RefList property's value.
' Value: Represents the list of cells that have use this validation formula.
' Remaraks: It has format of
'   "'WsName1'!Addr1,Addr2,...;'WsName2'!Addr1,Addr2,...;..." where WsName is
'   the Worksheet Name and Addr is the cell address. Only the first
'   REF_COUNT_LIMIT references are stored.
' Date Created: 2026-07-01
' Date Last Modified: 2026-07-01
'------------------------------------------------------------------------------'
Friend Property Get refList() As String
  refList = m_RefList
End Property
Public Property Let refList(ByVal value As String)
  m_RefList = value
End Property
