Option Explicit

'------------------------------------------------------------------------------'
' Class Module Name: ProtectionState
' Summary: Contains a worksheet's protection information.
' Date Created: 2026-05-23
' Date Last Modified: 2026-05-25
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Fields
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Constant Fields
'------------------------------------------------------------------------------'

' The name of this class module.
Private Const MODULE_NAME As String = "ProtectionState"

'------------------------------------------------------------------------------'
' Variable Fields
'------------------------------------------------------------------------------'

Private m_IsProtected As Boolean
Private m_ProtectContents As Boolean
Private m_ProtectDrawingObjects As Boolean
Private m_ProtectScenarios As Boolean
Private m_AllowFormattingCells As Boolean
Private m_AllowFormattingColumns As Boolean
Private m_AllowFormattingRows As Boolean
Private m_AllowInsertingColumns As Boolean
Private m_AllowInsertingRows As Boolean
Private m_AllowInsertingHyperlinks As Boolean
Private m_AllowDeletingColumns As Boolean
Private m_AllowDeletingRows As Boolean
Private m_AllowSorting As Boolean
Private m_AllowFiltering As Boolean
Private m_AllowUsingPivotTables As Boolean

'------------------------------------------------------------------------------'
' Pseudo-Constructors
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Summary: Instantiates and initializes an instance of the ProtectionState
'   class with protection state of the specified Worksheet.
' Parameter(s)
'   ws - The Worksheet.
' Returns: An instance of the ProtectionState class initalized with the
'   protection state of the specified Worksheet.
' Date Created: 2026-05-23
' Date Last Modified: 2026-05-25
'------------------------------------------------------------------------------'
Friend Sub Init(ByRef ws As Worksheet)
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "Init"

  m_IsProtected = ws.ProtectContents
  
  If m_IsProtected Then
    m_ProtectContents = ws.ProtectContents
    m_ProtectDrawingObjects = ws.ProtectDrawingObjects
    m_ProtectScenarios = ws.ProtectScenarios
      
    With ws.Protection
      m_AllowFormattingCells = .AllowFormattingCells
      m_AllowFormattingColumns = .AllowFormattingColumns
      m_AllowFormattingRows = .AllowFormattingRows
      m_AllowInsertingColumns = .AllowInsertingColumns
      m_AllowInsertingRows = .AllowInsertingRows
      m_AllowInsertingHyperlinks = .AllowInsertingHyperlinks
      m_AllowDeletingColumns = .AllowDeletingColumns
      m_AllowDeletingRows = .AllowDeletingRows
      m_AllowSorting = .AllowSorting
      m_AllowFiltering = .AllowFiltering
      m_AllowUsingPivotTables = .AllowUsingPivotTables
    End With
  End If
  
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
' Summary: Gets the IsProtected property's value.
' Value: The associated Worksheet's overall protection status.
' Date Created: 2026-05-03
' Date Last Modified: 2026-05-23
'------------------------------------------------------------------------------'
Friend Property Get IsProtected() As Boolean
  IsProtected = m_IsProtected
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the ProtectContents property's value.
' Value: The associated Worksheet's ProtectContents property.
' Date Created: 2026-05-03
' Date Last Modified: 2026-05-23
'------------------------------------------------------------------------------'
Friend Property Get ProtectContents() As Boolean
  ProtectContents = m_ProtectContents
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the ProtectDrawingObjects property's value.
' Value: The associated Worksheet's ProtectDrawingObjects property.
' Date Created: 2026-05-03
' Date Last Modified: 2026-05-23
'------------------------------------------------------------------------------'
Friend Property Get ProtectDrawingObjects() As Boolean
  ProtectDrawingObjects = m_ProtectDrawingObjects
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the ProtectScenarios property's value.
' Value: The associated Worksheet's ProtectScenarios property.
' Date Created: 2026-05-03
' Date Last Modified: 2026-05-23
'------------------------------------------------------------------------------'
Friend Property Get ProtectScenarios() As Boolean
  ProtectScenarios = m_ProtectScenarios
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the AllowFormattingCells property's value.
' Value: The associated Worksheet's Protection.AllowFormattingCells property.
' Date Created: 2026-05-03
' Date Last Modified: 2026-05-23
'------------------------------------------------------------------------------'
Friend Property Get AllowFormattingCells() As Boolean
  AllowFormattingCells = m_AllowFormattingCells
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the AllowFormattingColumns property's value.
' Value: The associated Worksheet's Protection.AllowFormattingColumns property.
' Date Created: 2026-05-03
' Date Last Modified: 2026-05-23
'------------------------------------------------------------------------------'
Friend Property Get AllowFormattingColumns() As Boolean
  AllowFormattingColumns = m_AllowFormattingColumns
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the AllowFormattingRows property's value.
' Value: The associated Worksheet's Protection.AllowFormattingRows property.
' Date Created: 2026-05-03
' Date Last Modified: 2026-05-23
'------------------------------------------------------------------------------'
Friend Property Get AllowFormattingRows() As Boolean
  AllowFormattingRows = m_AllowFormattingRows
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the AllowInsertingColumns property's value.
' Value: The associated Worksheet's Protection.AllowInsertingColumns property.
' Date Created: 2026-05-03
' Date Last Modified: 2026-05-23
'------------------------------------------------------------------------------'
Friend Property Get AllowInsertingColumns() As Boolean
  AllowInsertingColumns = m_AllowInsertingColumns
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the AllowInsertingRows property's value.
' Value: The associated Worksheet's Protection.AllowInsertingRows property.
' Date Created: 2026-05-03
' Date Last Modified: 2026-05-23
'------------------------------------------------------------------------------'
Friend Property Get AllowInsertingRows() As Boolean
  AllowInsertingRows = m_AllowInsertingRows
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the AllowInsertingHyperlinks property's value.
' Value: The associated Worksheet's Protection.AllowInsertingHyperlinks property.
' Date Created: 2026-05-03
' Date Last Modified: 2026-05-23
'------------------------------------------------------------------------------'
Friend Property Get AllowInsertingHyperlinks() As Boolean
  AllowInsertingHyperlinks = m_AllowInsertingHyperlinks
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the AllowDeletingColumns property's value.
' Value: The associated Worksheet's Protection.AllowDeletingColumns property.
' Date Created: 2026-05-03
' Date Last Modified: 2026-05-23
'------------------------------------------------------------------------------'
Friend Property Get AllowDeletingColumns() As Boolean
  AllowDeletingColumns = m_AllowDeletingColumns
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the AllowDeletingRows property's value.
' Value: The associated Worksheet's Protection.AllowDeletingRows property.
' Date Created: 2026-05-03
' Date Last Modified: 2026-05-23
'------------------------------------------------------------------------------'
Friend Property Get AllowDeletingRows() As Boolean
  AllowDeletingRows = m_AllowDeletingRows
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the AllowSorting property's value.
' Value: The associated Worksheet's Protection.AllowSorting property.
' Date Created: 2026-05-03
' Date Last Modified: 2026-05-23
'------------------------------------------------------------------------------'
Friend Property Get AllowSorting() As Boolean
  AllowSorting = m_AllowSorting
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the AllowFiltering property's value.
' Value: The associated Worksheet's Protection.AllowFiltering property.
' Date Created: 2026-05-03
' Date Last Modified: 2026-05-23
'------------------------------------------------------------------------------'
Friend Property Get AllowFiltering() As Boolean
  AllowFiltering = m_AllowFiltering
End Property

'------------------------------------------------------------------------------'
' Summary: Gets the AllowUsingPivotTables property's value.
' Value: The associated Worksheet's Protection.AllowUsingPivotTables property.
' Date Created: 2026-05-03
' Date Last Modified: 2026-05-23
'------------------------------------------------------------------------------'
Friend Property Get AllowUsingPivotTables() As Boolean
  AllowUsingPivotTables = m_AllowUsingPivotTables
End Property



'------------------------------------------------------------------------------'
' Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Friend Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Summary: Restores the original protection status to a target Worksheet.
' Parameters:
'   ws - The target Worksheet object to protect.
'   state - The read-only clsProtectionState object holding original values.
'   password: Optional string password required to lock the sheet.
' Date Created: 2026-05-23
' Date Last Modified: 2026-05-23
'------------------------------------------------------------------------------'
Friend Sub RestoreProtection( _
  ByRef ws As Worksheet, _
  Optional ByVal Password As String = "")
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RestoreProtection"
  
  ' Guard clauses to ensure data objects exist and that the sheet was
  ' originally protected.
  If ws Is Nothing Then GoTo Exit_Proc
  
  If Not m_IsProtected Then GoTo Exit_Proc
    
  ' Re-apply all original restrictions reading from the class properties
  ws.Protect Password:=Password, _
    Contents:=m_ProtectContents, _
    DrawingObjects:=m_ProtectDrawingObjects, _
    Scenarios:=m_ProtectScenarios, _
    AllowFormattingCells:=m_AllowFormattingCells, _
    AllowFormattingColumns:=m_AllowFormattingColumns, _
    AllowFormattingRows:=m_AllowFormattingRows, _
    AllowInsertingColumns:=m_AllowInsertingColumns, _
    AllowInsertingRows:=m_AllowInsertingRows, _
    AllowInsertingHyperlinks:=m_AllowInsertingHyperlinks, _
    AllowDeletingColumns:=m_AllowDeletingColumns, _
    AllowDeletingRows:=m_AllowDeletingRows, _
    AllowSorting:=m_AllowSorting, _
    AllowFiltering:=m_AllowFiltering, _
    AllowUsingPivotTables:=m_AllowUsingPivotTables
               
Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

