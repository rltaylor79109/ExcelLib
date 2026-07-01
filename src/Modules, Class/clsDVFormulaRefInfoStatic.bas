Option Explicit

'------------------------------------------------------------------------------'
' Class Module Name: clsDVFormulaRefInfoStatic
' Summary: It contains members that are intended to be used as pseudo-static
'   members of the clsDVFormulaRefInfoStatic class.
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
Private Const MODULE_NAME As String = "clsDVFormulaRefInfoStatic"

'------------------------------------------------------------------------------'
' Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Friend Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Summary: Instantiates and initializes an instance of the
'   clsDVFormulaRefInfo class with the specified values. It is a factory method.
' Parameter(s)
'   ws - The Worksheet.
' Return(s): An instance of the clsDVFormulaRefInfoStatic class initalized
'   with the specified property values.
' Date Created: 2026-07-01
' Date Last Modified: 2026-07-01
'------------------------------------------------------------------------------'
Friend Function Create( _
  ByVal pFormula1 As String, _
  ByVal pFormula2 As String, _
  Optional ByVal pRefCount As Long = 0, _
  Optional ByRef pRefCountExceedsLimit As Boolean = False, _
  Optional ByRef pRefList As String = "" _
  ) As clsDVFormulaRefInfoStatic
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "Create"
  
  Dim instance As clsDVFormulaRefInfo
  Set instance = New clsDVFormulaRefInfo
  instance.Init _
    formula1:=pFormula1, _
    formula2:=pFormula2, _
    refCount:=pRefCount, _
    refCountExceedsLimit:=pRefCountExceedsLimit, _
    refList:=pRefList
    
  Set Create = instance

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

