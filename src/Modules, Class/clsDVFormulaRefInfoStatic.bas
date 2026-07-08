Option Explicit

'------------------------------------------------------------------------------'
' Class Module Name: clsDVFormulaRefInfoStatic
' Summary: It contains members that are intended to be used as pseudo-static
'   members of the clsDVFormulaRefInfoStatic class.
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
Private Const MODULE_NAME As String = "clsDVFormulaRefInfoStatic"

'------------------------------------------------------------------------------'
' Properties
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Friend Methods
'------------------------------------------------------------------------------'

''------------------------------------------------------------------------------'
'' Summary: Instantiates and initializes an instance of the
''   clsDVFormulaRefInfo class initialized with the specified property values.
'' Parameter(s)
''   pFormula1 - Represents the cell's Validation Formula1 value.
''   pFormula2 - Represents the cell's Validation Formula1 value. A value of ""
''     indicates its value is undefined. It is only defined when the validation
''     Type property is set to xlValidateList or xlValidateCustom and the
''     validation Operator is set either xlBetween or xlNotBetween.
''   pRefDict - Represents the list of cells that have use this validation
''     formula. It contains a dictionary of collections of cell addresses
''     keyed by their Worksheet names. It is an optional parameter with a
''     default value of nothing.
'' Return(s): An instance of the clsDVFormulaRefInfoStatic class initalized
''   with the specified property values.
'' Date Created: 2026-07-02
'' Date Last Modified: 2026-07-07
''------------------------------------------------------------------------------'
'Friend Function Create( _
'  ByVal pF1 As String, _
'  ByVal pF2 As String, _
'  Optional ByRef pRefDict As Dictionary = Nothing _
'  ) As clsDVFormulaRefInfo
'
'  On Error GoTo Err_Proc
'  Const METHOD_NAME As String = "Create"
'
'  Dim instance As clsDVFormulaRefInfo
'  Set instance = New clsDVFormulaRefInfo
'  instance.Init _
'    pF1:=pF1, _
'    pF2:=pF2, _
'    pRefDict:=pRefDict
'
'  Set Create = instance
'
'Exit_Proc:
'  Exit Function
'Err_Proc:
'  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
'  Resume Exit_Proc
'End Function

'------------------------------------------------------------------------------'
' Summary: Instantiates and initializes an instance of the
'   clsDVFormulaRefInfo class with the data validation information from the
'   specified cell.
' Parameter(s)
'   pCell - The cell whose validation information is used.
'   pRefDict - Represents the list of cells that have use this validation
'     formula. It contains a dictionary of collections of cell addresses
'     keyed by their Worksheet names.It is an optional parameter with a
'     default value of nothing.
' Return(s): An instance of the clsDVFormulaRefInfoStatic class initalized
'   with the specified property values.
' Date Created: 2026-07-07
' Date Last Modified: 2026-07-08
'------------------------------------------------------------------------------'
Friend Function CreateFromCell( _
  ByRef pCell As Range, _
  Optional ByRef pRefDict As Dictionary = Nothing _
  ) As clsDVFormulaRefInfo
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "CreateFromCell"
  
  Dim instance As clsDVFormulaRefInfo
  Set instance = New clsDVFormulaRefInfo
  instance.InitWithCell _
    pCell:=pCell, _
    pRefDict:=pRefDict
  Set CreateFromCell = instance
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function
