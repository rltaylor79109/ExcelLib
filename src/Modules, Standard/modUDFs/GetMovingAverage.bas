'------------------------------------------------------------------------------'
' Summary: Gets the moving average of the specified period of the values
' contained in the specified range.
' Parameter(s):
'   pRange - Range in which to find the moving average.
'   period - The period of the moving average.
'   notFoundValue - The value to return if pRange is contains no non-error
'     values. It is an optional parameter with a default value of
'     CVErr(xlErrNA). This is the same CVErr(2042) and as the value returned
'     by the NA() spreadsheet function.
' Return: The moving average of the specified period of the values contained
'   in the specified range.
' Date Created: 2017-03-24
' Date Last Modified: 2026-05-19
'------------------------------------------------------------------------------'
Public Function GetMovingAverage( _
  pRange As Range, _
  period As Integer, _
  Optional notFoundValue As Variant _
  ) As Double
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "GetMovingAverage"

  Dim count As Integer
  Dim i As Integer
  Dim sum As Double
       
  On Error GoTo Err_Proc
  
  If IsMissing(notFoundValue) Then
    notFoundValue = CVErr(xlErrNA)
  End If

  count = 0
  i = pRange.count
  sum = 0
  
  ' Sum values that are not missing until the number indicated by the argument,
  ' intPeriod, have been averaged or until we reach the top of the column, whichever
  ' comes first.
  Do While count < period And i > 0
    ' VBA evaluates both parameters in an OR statement so if we use IsError(X)
    ' or Len(X) > 0 and X evalautes to an error, an error will be raised.
    ' Consequentailly we must use nested ifs here.
    If Not IsError(pRange.Item(i)) Then
      If Len(pRange.Item(i)) > 0 Then
        count = count + 1
        sum = sum + pRange.Item(i)
      End If
    End If
    i = i - 1
  Loop
  
  If count <> 0 Then
    GetMovingAverage = sum / count
  Else
    ' This could happen if the column is completely empty
    ' It is undefined because there is no data in the range
    GetMovingAverage = 0
  End If

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, mModuleName, METHOD_NAME
  Resume Exit_Proc
End Function