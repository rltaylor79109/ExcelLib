'------------------------------------------------------------------------------'
' Summary: Sets the horizontal aligment format of the current selection to
'   center across selection.
' Remarks: Did not work properly until the method cleared the center across
'   selection for before applying it.
' Date Created: 2026-05-14
' Date Last Modified: 2026-08-02
'------------------------------------------------------------------------------'
Public Sub CenterAcrossSelecton()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "CenterAcrossSelecton"
  
  Dim cell As Range
  Dim errorOccurred As Boolean
  Dim targetRng As Range

  errorOccurred = False
  Application.ScreenUpdating = False
  
  ' Check if user selected cells
  If TypeName(Selection) <> "Range" Then GoTo Exit_Proc
   
  Set targetRng = Selection
  
  ' Step 1: Clear xlCenterAcrossSelection ONLY within your current selection
  ' to break any lingering connections from previous formatting runs
  For Each cell In targetRng
      If cell.HorizontalAlignment = xlCenterAcrossSelection Then
        cell.HorizontalAlignment = xlGeneral
      End If
  Next cell
  
  ' Step 2: Apply Center Across Selection to the selected 4-column block
  targetRng.HorizontalAlignment = xlCenterAcrossSelection

Exit_Proc:
  Application.ScreenUpdating = True
  
  If errorOccurred Then
    ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  End If
  
  Exit Sub
Err_Proc:
  errorOccurred = True
  Resume Exit_Proc
End Sub