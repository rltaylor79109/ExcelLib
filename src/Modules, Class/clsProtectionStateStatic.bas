Option Explicit

'------------------------------------------------------------------------------'
' Class Module Name: clsProtectionStateStatic
' Summary: It contains members that are intended to be used as pseudo-static
'   members of the clsProtectionState class.
' Date Created: 2026-05-23
' Date Last Modified: 2026-08-16
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Fields
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Constant Fields
'------------------------------------------------------------------------------'

' The name of this class module.
Private Const MODULE_NAME As String = "clsProtectionStateStatic"

'------------------------------------------------------------------------------'
' Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Friend Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Summary: Instantiates and initializes an instance of the clsProtectionState
'   class with the protection state of the specified Worksheet. It is a
'   factory method.
' Parameter(s)
'   ws - The Worksheet.
' Return(s): An instance of the clsProtectionState class initalized with the
'   protection state of the specified Worksheet.
' Date Created: 2026-05-25
' Date Last Modified: 2026-08-16
'------------------------------------------------------------------------------'
Friend Function Create(ByVal ws As Worksheet) As clsProtectionState
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "Create"
  
  ' Guard against unassigned worksheet references
  If ws Is Nothing Then Exit Function

  Set Create = New clsProtectionState
  Create.Init ws
    
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Instantiates and initializes an instance of the clsProtectionState
'   class with the protection state of the specified Worksheet. It then
'   unprotects the specified Worksheet. It is a factory method.
' Parameter(s)
'   ws - The Worksheet.
' Date Created: 2026-07-22
' Date Last Modified: 2026-07-22
'------------------------------------------------------------------------------'
Friend Function CreateAndUnprotect( _
  ByRef ws As Worksheet _
  ) As clsProtectionState
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "SaveAndUnprotect"

  Set CreateAndUnprotect = Create(ws)
  ws.Unprotect
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function
Option Explicit

'------------------------------------------------------------------------------'
' Class Module Name: clsProtectionStateStatic
' Summary: It contains members that are intended to be used as pseudo-static
'   members of the clsProtectionState class.
' Date Created: 2026-05-23
' Date Last Modified: 2026-07-01
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Fields
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Constant Fields
'------------------------------------------------------------------------------'

' The name of this class module.
Private Const MODULE_NAME As String = "clsProtectionStateStatic"

'------------------------------------------------------------------------------'
' Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Friend Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Summary: Instantiates and initializes an instance of the clsProtectionState
'   class with the protection state of the specified Worksheet. It is a
'   factory method.
' Parameter(s)
'   ws - The Worksheet.
' Return(s): An instance of the clsProtectionState class initalized with the
'   protection state of the specified Worksheet.
' Date Created: 2026-05-25
' Date Last Modified: 2026-05-25
'------------------------------------------------------------------------------'
Friend Function Create(ws As Worksheet) As clsProtectionState
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "Create"
  
  Dim ps As clsProtectionState
  Set ps = New clsProtectionState
  ps.Init ws
  Set Create = ps

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function
