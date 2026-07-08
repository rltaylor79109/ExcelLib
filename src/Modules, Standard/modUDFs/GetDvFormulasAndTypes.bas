'------------------------------------------------------------------------------'
' Summary: Gets the data validation Formula1 and Forumla2 properties and their
'   types for the specified cell.
' Parameter(s):
'   targetCell - The cell whose data validation properties are returned.
'   op - The validation formula operator. A value of 0 represents undefined.
'   f1 - The value of the data validation Formula1 property of the specified
'     cell. It is an output parameter.
'   f1Type - The data validation Formula1's type. It is an output parameter.
'   f2 - The value of the data validation Formula2 property of the specified
'     cell. It is an output parameter.
'   f2Type - The data validation Formula2's type. It is an output parameter.
' Date Created: 2026-07-05
' Date Last Modified: 2026-07-07
'------------------------------------------------------------------------------'
Public Sub GetDvFormulasAndTypes( _
  targetCell As Range, _
  ByRef op As XlFormatConditionOperator, _
  ByRef f1 As String, _
  ByRef f1Type As ValidationSrcType, _
  ByRef f2 As String, _
  ByRef f2Type As ValidationSrcType _
  )

  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "GetDvFormulasAndTypes"
  
  Const VAL_OP_IS_UNDEFINED = 0
  ' Actual function is VBA user defined function, Excel built-in
  ' function, or LAMDA function.
  Const REGEX_PATTERN_TO_MATCH_ACTUAL_FUNCTION As String = _
    "^[A-Z0-9._]+(?=\()"
  Const REGEX_PATTERN_TO_MATCH_STANDARD_CELL_REF As String = _
    "^(?:.+!)?\$?[A-Z]+\$?[0-9]+#?$"
    
  Dim f2IsDefined As Boolean
  Dim regX As RegExp
  Dim nm As Name
  
  ' Get the operator.
  ' Check the type first
  Select Case targetCell.Validation.Type
    ' Cases where it is safe and meaningful to evaluate the operator.
    Case _
        xlValidateWholeNumber, _
        xlValidateDecimal, _
        xlValidateDate, _
        xlValidateTime, _
        xlValidateTextLength
         
        ' It is safe and meaningful to evaluate the operator here
        op = targetCell.Validation.Operator
    ' Cases where the operator property is ignore. We set the value to
    ' 0 to represent undefined.
    Case _
      xlValidateInputOnly, _
      xlValidateList, _
      xlValidateCustom
      op = VAL_OP_IS_UNDEFINED
  End Select
  
  
  ' 1. Does the cell have validation?
  On Error Resume Next
  f1 = targetCell.Validation.formula1
  err.Clear
  On Error GoTo Err_Proc
  If f1 = "" Then
    f1Type = vstNoValidation
    f2 = ""
    f2Type = vstNoValidation
    GoTo Eval_F2
  End If
  
  ' 2. If it doesn't start with "=", it's a hardcoded list
  ' (e.g., "Yes,No,Maybe")
  If Left(f1, 1) <> "=" Then
    f1Type = vstHardCodedValue
    GoTo Eval_F2
  End If

  ' Strip the leading "=" for analysis
  f1 = Mid(f1, 2)
  
  ' 3. Is it a standard cell reference (e.g., $A$1, $B$5:$B$10)
  Set regX = New RegExp
  regX.Pattern = REGEX_PATTERN_TO_MATCH_STANDARD_CELL_REF
  regX.IgnoreCase = True
  If regX.test(f1) Then
    f1Type = vstCellRef
    GoTo Eval_F2
  End If
    
  ' 4. Is it a Named Range?
  On Error Resume Next
  Set nm = targetCell.Worksheet.Parent.Names(f1) ' Workbook level
  If nm Is Nothing Then
    Set nm = targetCell.Worksheet.Names(f1) ' Worksheet level
  End If
  err.Clear
  On Error GoTo Err_Proc
    
  If Not nm Is Nothing Then
    f1Type = vstDefinedName
    GoTo Eval_F2
  End If
    
  ' 5. Is it a function (Built-in, UDF, or LAMBDA)
  regX.Pattern = REGEX_PATTERN_TO_MATCH_ACTUAL_FUNCTION
  If regX.test(f1) Or InStr(1, f1, "LAMBDA", vbTextCompare) > 0 Then
    f1Type = vstFunction
    GoTo Eval_F2
  End If
    
  ' Fallback case
  f1Type = vstUndefined
  
Eval_F2:
  ' If f1 is undefined or is empty, then f2 is undefined.
  If f1Type = vstUndefined Or f1Type = vstNoValidation Then
    f2 = ""
    f2Type = vstUndefined
    GoTo Exit_Proc
  End If
  
  f2 = ""
  On Error Resume Next
  f2 = targetCell.Validation.formula2
  err.Clear
  On Error GoTo Err_Proc
  If f2 = "" Then
    f2IsDefined = False
    f2Type = vstUndefined
    GoTo Exit_Proc
  End If
  
  ' 2. If it doesn't start with "=", it's a hardcoded list
  ' (e.g., "Yes,No,Maybe")
  If Left(f2, 1) <> "=" Then
    f2Type = vstHardCodedValue
    GoTo Exit_Proc
  End If

  ' Strip the leading "=" for analysis
  f2 = Mid(f2, 2)
  
  ' 3. Is it a standard cell reference (e.g., $A$1, $B$5:$B$10)
  Set regX = New RegExp
  regX.Pattern = REGEX_PATTERN_TO_MATCH_STANDARD_CELL_REF
  regX.IgnoreCase = True
  If regX.test(f2) Then
    f2Type = vstCellRef
    GoTo Exit_Proc
  End If
    
  ' 4. Is it a Named Range?
  On Error Resume Next
  Set nm = targetCell.Worksheet.Parent.Names(f2) ' Workbook level
  If nm Is Nothing Then
    Set nm = targetCell.Worksheet.Names(f2) ' Worksheet level
  End If
  err.Clear
  On Error GoTo Err_Proc
    
  If Not nm Is Nothing Then
    f2Type = vstDefinedName
    GoTo Exit_Proc
  End If
    
  ' 5. Is it a function (Built-in, UDF, or LAMBDA)
  regX.Pattern = REGEX_PATTERN_TO_MATCH_ACTUAL_FUNCTION
  If regX.test(f2) Or InStr(1, f2, "LAMBDA", vbTextCompare) > 0 Then
    f1Type = vstFunction
    GoTo Exit_Proc
  End If
    
  ' Fallback case
  f2Type = vstUndefined

Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub