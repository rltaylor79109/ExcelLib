'------------------------------------------------------------------------------'
' Summary: Determines if a worksheet is a reference report worksheets using
'   the specified worksheet CodeName property.
' Remarks: It is usually used to avoid searching the worksheets for defined
'   name or table list object references.
' Parameter(s):
'   pCodeName - The Codename property of the worksheet to test.
' Return(s): True if the worksheet specified by its Codename is a reference
'   report worksheet; otherwise False.
' Date Created: 2026-07-12
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Private Function IsWsRefRpt(pCodeName As String) As Boolean
  Const METHOD_NAME = "IsWsRefRpt"
  
  Static dict As Dictionary
  If dict Is Nothing Then
    Set dict = GetRefRptCodeNameList()
  End If

  IsWsRefRpt = dict.Exists(pCodeName)
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function