'------------------------------------------------------------------------------'
' Summary: Gets dictionary that represents a list of the code names of the
'   reference report worksheets.
' Remarks: It is usually used to avoid searching the worksheets for defined
'   name or table list object references.
' Return(s): A dictionary that represents a list of the code names of the
'   reference report worksheets.
' Date Created: 2026-07-12
' Date Last Modified: 2026-07-13
'------------------------------------------------------------------------------'
Private Function GetRefRptCodeNameList() As Dictionary
  Const METHOD_NAME = "GetRefRptCodeNameList"
  
  Dim dict As Dictionary
  Set dict = New Dictionary
  
  dict.Add key:="SheetDefNamesAllRefs", item:=True
  dict.Add key:="SheetDefNamesWsRefsRpt", item:=True
  dict.Add key:="SheetDvAllRpt", item:=True
  dict.Add key:="SheetDvByForumlaRpt", item:=True
  dict.Add key:="SheetTblsAllRefsRpt", item:=True
  dict.Add key:="SheetTblsWsRefsRpt", item:=True

  Set GetRefRptCodeNameList = dict
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function