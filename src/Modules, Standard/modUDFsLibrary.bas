Option Explicit

'------------------------------------------------------------------------------'
' Module Name: modUDFsLibrary
' Summary: Contains library (not application specific) user defined functions.
' Date Created: 2017-03-24
' Date Last Modified: 2026-08-31
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Fields
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Constant Fields
'------------------------------------------------------------------------------'

' The name of this module.
Private Const MODULE_NAME As String = "modUDFsLibrary"

'------------------------------------------------------------------------------'
' Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Public Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Summary: Determines if the specified VBA component contains the specified
'   string.
' Parameter(s):
'   searchIn - The VBA component (Modules, Sheets, Forms, Classes) to search
'     in.
'   searchFor - The string to search for.
'   wholeWordOnly - If True, the search matches whole words only, so that
'     "Pie" matches "Apple Pie" but does not match "Pied Piper"; otherwise,
'     the match occurs whether the string is within a word or not
'     so that "Pie" matches both "Apple Pie" and "Pied Piper". It is an
'     optional parameter with a default value of False.
'   matchCase - If True, the search is case-sensitive; otherwise the
'     search is case-insensitive. It is an optional parameter with a default
'     value of False.
' Returns: True if the specified VBA component contains the specified
'   string; otherwise False.
' Date Created: 2026-06-16
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Function ComponentContainsString( _
  searchIn As VBComponent, _
  searchFor As String, _
  Optional wholeWordOnly As Boolean = False, _
  Optional matchCase As Boolean = False _
  ) As Boolean

  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ComponentContainsString"
  Dim codeMod As codeModule
  Dim i As Long
  Dim lineText As String
  ComponentContainsString = False
  
  Set codeMod = searchIn.codeModule
  For i = 1 To codeMod.CountOfLines
    lineText = codeMod.Lines(i, 1)
    If ContainsString( _
      searchIn:=lineText, _
      searchFor:=searchFor, _
      wholeWordOnly:=wholeWordOnly, _
      matchCase:=matchCase) Then

      ComponentContainsString = True
      Exit For
    End If
  Next i

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Determines if the specified string contains any characters
'   contained in another string.
' Parameter(s):
'   pText - The string searched to determine if it contains any of the
'     characters contained by the charsToFind string.
'   charsToFind - A string that represents the list of characters to find
'     in the pText parameter.
' Returns: True if the pText parameter string contains any characters
'   contained charsToFind parameter; otherwise, False.
' Date Created: 2026-05-18
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Function ContainsAnyChar(pText As String, charsToFind As String) As Boolean
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ContainsAnyChar"

  Dim i As Long
  Dim curChar As String
    
  ' Loop through each character in the "charsToFind" string
  For i = 1 To Len(charsToFind)
    curChar = Mid(charsToFind, i, 1)
        
  ' If the character is found anywhere in the text, return True and exit
  If InStr(1, pText, curChar, vbTextCompare) > 0 Then
    ContainsAnyChar = True
    GoTo Exit_Proc
    End If
  Next i

  ContainsAnyChar = False
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Determines whether a string contains any of the characters that
'   are not permitted in sheet names. The prohibited characters are
'   \ / * [ ] ? :
' Parameter(s):
'   pSheetName - The string to test.
' Return:   True if pSheetName contains any of the prohibited characters,
'   otherwise false.
' Date Created: 2018-08-30
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Function ContainsIllegalWsNameChars( _
  pSheetName As String) As Boolean
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ContainsIllegalWsNameChars"
  
  ContainsIllegalWsNameChars = _
    ContainsAnyChar(pText:=pSheetName, _
    charsToFind:=WS_NAME_FORBIDDEN_CHARS)

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Determines if the specified string contains any leading or
'   trailing spaces.
' Parameter(s):
'   pText - The string to test.
' Return:   True if pText contains any leading spaces, trailing spaces, or
'   both; otherwise false.
' Date Created: 2026-05-18
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Function ContainsLeadingOrTrailingSpaces(pText As String) As Boolean
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ContainsLeadingOrTrailingSpaces"
  
  ContainsLeadingOrTrailingSpaces = pText <> Trim(pText)

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Determines if one string contains another string search for whole
'   words only.
' Parameter(s):
'   searchIn - The main string you want to search inside of.
'   searchFor - The string you are looking for.
'   wholeWordOnly - If True, the search matches whole words only, so that
'     "Pie" matches "Apple Pie" but does not match "Pied Piper"; otherwise,
'     the match occurs whether the string is within a word or not
'     so that "Pie" matches both "Apple Pie" and "Pied Piper". It is an
'     optional parameter with a default value of False.
'   matchCase - If True, the search is case-sensitive; otherwise the
'     search is case-insensitive. It is an optional parameter with a default
'     value of False.
' Returns: True if searchIn contains the whole wordsearchFor; otherwise False.
' Date Created: 2026-06-15
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Function ContainsString( _
  ByVal searchIn As String, _
  ByVal searchFor As String, _
  Optional ByVal wholeWordOnly As Boolean = False, _
  Optional ByVal matchCase As Boolean = False _
  ) As Boolean
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ContainsString"
  
  Dim escapedSearchFor As String
  Dim finalPattern As String
  Dim RegEx As RegExp

  Set RegEx = New RegExp
  
  ' Quick exit if either string is empty
  If Len(searchIn) = 0 Or Len(searchFor) = 0 Then
      ContainsString = False
      GoTo Exit_Proc
  End If
  
  escapedSearchFor = EscapeRegexChars(searchFor)
  
  If wholeWordOnly Then
      finalPattern = "\b" & escapedSearchFor & "\b"
  Else
      finalPattern = escapedSearchFor
  End If
  
  With RegEx
      .Pattern = finalPattern
      .IgnoreCase = Not matchCase
      .Global = False
  End With
  
  ContainsString = RegEx.test(searchIn)
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Escapes characters that have special meanings in RegEx in the
'   specified string.
' Parameter(s):
'   pString - The string to modify.
' Returns: The input string modified so that characters that have special
'   meanings in RegEx are escaped.
' Date Created: 2026-06-15
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Function EscapeRegexChars(ByVal pString As String) As String
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "EscapeRegexChars"
  
  Const SPECIAL_CHARACTERS As String = "\^$.|?*+()[]{}"
  
  Dim char As String
  Dim escapedStr As String
  Dim i As Long

  escapedStr = pString
  For i = 1 To Len(SPECIAL_CHARACTERS)
    char = Mid(SPECIAL_CHARACTERS, i, 1)
    escapedStr = Replace(escapedStr, char, "\" & char)
  Next i
  EscapeRegexChars = escapedStr
    
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Gets a chartsheet object by its CodeName property.
' Remarks: A Chartsheet's CodeName property is distinct from its Name property.
'   The CodeName property is visible in the Project Explorer in the Excel
'   VBA IDE. The Name property is shown on the Chart tab in the regular
'   Excel window. The method searches the active workbook.
' Parameter(s):
'   pCodeName - Represents the value of the CodeName property of the chart of
'     interest.
' Return(s): The chartsheet object with the specified CodeName property.
' Date Created: 2026-05-17
' Date Last Modified: 2026-08-13
'------------------------------------------------------------------------------'
Public Function GetCsByCodeName(pCodeName As String) As Chart
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "GetCsByCodeName"
  
  Dim errorOccurred As Boolean
  Dim cs As Chart
  Dim wb As Workbook
  
  OptimizeAppEnvForSpeed True
  errorOccurred = False
  Set wb = ThisWorkbook
  For Each cs In wb.Charts
    If UCase(cs.codeName) = UCase(pCodeName) Then
      Set GetCsByCodeName = cs
      GoTo Exit_Proc
    End If
  Next cs
  
Exit_Proc:
  OptimizeAppEnvForSpeed False
  If errorOccurred Then
    ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  End If
  Exit Function
Err_Proc:
  errorOccurred = True
  Resume Exit_Proc
End Function

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
' Date Last Modified: 2026-07-12
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
  Dim nm As name
  
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
  f1 = targetCell.Validation.Formula1
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

'------------------------------------------------------------------------------'
' Summary: Safely retrieves the Range object that a global defined name refers
'   to.
' Remarks: Checks if the defined name refers to a valid Range vs. a constant,
'   formula, or broken reference.
' Parameter(s):
'   pDefNameStr - A string that represents the defined name.
' Return(s): If the defined name refers to a cell range, the range object
'   that a global defined name refers to. Otherwise, Nothing if it refers to
'   a constant/formula, invalid reference, or if the defined name does not
'   exist.
' Date Created: 2026-08-08
' Date Last Modified: 2026-08-08
'------------------------------------------------------------------------------'
Public Function GetDefNameRange(ByVal pDefNameStr As String) As Range
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "GetDefNameRange"
    
  Dim nm As name
  Dim targetRng As Range
  
  ' Try to fetch the defined name
  On Error Resume Next
  Set nm = ThisWorkbook.Names(pDefNameStr)
  err.Clear
  On Error GoTo Err_Proc
  
  ' If the defined name doesn't exist, return Nothing
  If nm Is Nothing Then
    Set GetDefNameRange = Nothing
    GoTo Exit_Proc
  End If
  
  ' Check if the defined name resolves to a Range object.
  ' Note: .RefersToRange throws an error (1004) if the name refers to
  ' a constant, non-range formula, or broken reference (#REF!).
  On Error Resume Next
  Set targetRng = nm.RefersToRange
  err.Clear
  On Error GoTo Err_Proc
  
  ' If targetRng is valid, return the Range; otherwise return Nothing
  If Not targetRng Is Nothing Then
    Set GetDefNameRange = targetRng
  Else
    Set GetDefNameRange = Nothing
  End If
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Safely retrieves the value of any global defined name
' (whether it refers to a Range or a direct Constant/Formula).
' Remarks: This avoids errors from active sheet changes or duplicate locally
'   scoped defined names.
' Parameter(s):
'   pDefNameStrStr - A string that represents the the defined name.
' Return(s): The worksheet object with the specifed CodeName property.
' Date Created: 2026-07-22
' Date Last Modified: 2026-08-10
'------------------------------------------------------------------------------'
Public Function GetDefNameVal(ByVal pDefNameStr As String) As Variant
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "GetDefNameVal"
    
  Dim nm As name
  Dim targetRng As Range
  
  On Error Resume Next
  Set nm = ThisWorkbook.Names(pDefNameStr)
  err.Clear
  On Error GoTo Err_Proc
  
  ' If the name doesn't exist, return Empty/Error
  If nm Is Nothing Then
    GetDefNameVal = Empty
    GoTo Exit_Proc
  End If
  
  ' Check if the name actually points to a worksheet Range
  On Error Resume Next
  Set targetRng = nm.RefersToRange
  err.Clear
  On Error GoTo Err_Proc
  
  If Not targetRng Is Nothing Then
    ' It's a range reference (e.g., refers to $C$1)
    GetDefNameVal = targetRng.value
  Else
    ' It's a constant or formula (e.g., refers to =7)
    GetDefNameVal = Application.Evaluate(nm.RefersTo)
  End If
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Gets a worksheet object by its CodeName property.
' Remarks: A Worksheet's CodeName property is distinct from its Name property.
'   The CodeName property is visible in the Project Explorer in the Excel
'   VBA IDE. The Name property is shown on the Worksheet tab in the regular
'   Excel window. The method searches the active workbook.
' Parameter(s):
'   pCodeName - Represents the value of the CodeName property of the sheet of
'     interest.
' Return(s): The worksheet object with the specifed CodeName property.
' Date Created: 2026-05-17
' Date Last Modified: 2026-08-20
'------------------------------------------------------------------------------'
Public Function GetWsByCodeName(pCodeName As String) As Worksheet
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "GetWsByCodeName"
  
  Dim sheetFound As Boolean
  Dim wb As Workbook
  Dim ws As Worksheet
  
  Set wb = ThisWorkbook
  
  sheetFound = False
  For Each ws In wb.Worksheets
      If UCase(ws.codeName) = UCase(pCodeName) Then
        Set GetWsByCodeName = ws
        sheetFound = True
        GoTo Exit_Proc
      End If
  Next ws
  
  If Not sheetFound Then
    Set GetWsByCodeName = Nothing
  End If
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Gets the Name property from the Worksheet specified by CodeName.
' Remarks: A Worksheet's CodeName property is distinct from its Name property.
'   The CodeName property is visible in the Project Explorer in the Excel
'   VBA IDE. The Name property is shown on the Worksheet tab in the regular
'   Excel window. The method searches the active workbook.
' Parameter(s):
'   pCodeName - Represents the value of the CodeName property of the sheet of
'     interest.
' Return(s): The value of the Name property from the Worksheet specified by
  ' CodeName.
' Date Created: 2026-05-29
' Date Last Modified: 2026-08-20
'------------------------------------------------------------------------------'
Public Function GetWsNameByCodeName(pCodeName As String) As String
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "GetWsNameByCodeName"
  
  Dim ws As Worksheet
  
  Set ws = GetWsByCodeName(pCodeName)
  If ws Is Nothing Then
    GetWsNameByCodeName = "NOT FOUND"
  Else
    GetWsNameByCodeName = ws.name
  End If
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Determines if the specified cell has Data Validation.
' Parameter(s):
'   cell - Represents the target cell Range.
' Return(s): True if the specified cell has Data Validation; otherwise, False.
' Date Created: 2026-08-02
' Date Last Modified: 2026-08-02
'------------------------------------------------------------------------------'
Public Function HasValidation(cell As Range) As Boolean
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "HasValidation"
  
  On Error Resume Next
  HasValidation = (cell.Validation.Type >= 0)
  err.Clear
  On Error GoTo Err_Proc
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Spells out a whole number. For example, 152 returns
'   "One Hundred Fifty-Two".
' Parameter(s):
'   pNumber - The number
' Return(s): The spelled out number.
' Date Created: 2026-07-23
' Date Last Modified: 2026-07-23
'------------------------------------------------------------------------------'
Public Function SpellNumber(ByVal pNumber As Long) As String
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "SpellNumber"
  
  Dim chunkStr As String
  Dim chunkValue As Integer
  Dim count As Integer
  Dim decimalPlaces As Integer
  Dim hundredsStr As String
  Dim isNegative As Boolean
  Dim scaleunits As Variant
  Dim tempStr As String
  Dim tens As Variant
  Dim units As Variant
  
   ' Handle zero
   If pNumber = 0 Then
       SpellNumber = "Zero"
       GoTo Exit_Proc
   End If

  ' 2. Handle negative numbers
  If pNumber < 0 Then
      isNegative = True
      pNumber = Abs(CDbl(pNumber))
  End If
    
    units = Array( _
      "", _
      "One", _
      "Two", _
      "Three", _
      "Four", _
      "Five", _
      "Six", _
      "Seven", _
      "Eight", _
      "Nine", _
      "Ten", _
      "Eleven", _
      "Twelve", _
      "Thirteen", _
      "Fourteen", _
      "Fifteen", _
      "Sixteen", _
      "Seventeen", _
      "Eighteen", _
      "Nineteen")
    tens = Array( _
      "", _
      "", _
      "Twenty", _
      "Thirty", _
      "Forty", _
      "Fifty", _
      "Sixty", _
      "Seventy", _
      "Eighty", _
      "Ninety")
    scaleunits = Array( _
      "", _
      "Thousand", _
      "Million", _
      "Billion", _
      "Trillion")
    
    tempStr = ""
    count = 0
    
    Do While pNumber > 0
      chunkValue = pNumber Mod 1000
      If chunkValue > 0 Then
        hundredsStr = ""
        
        ' Hundreds
        If chunkValue \ 100 > 0 Then
          hundredsStr = units(chunkValue \ 100) & " Hundred "
        End If
        
        ' tens & units
        Select Case chunkValue Mod 100
          Case 1 To 19
              hundredsStr = hundredsStr & units(chunkValue Mod 100)
          Case Is >= 20
            hundredsStr = hundredsStr & tens((chunkValue Mod 100) \ 10)
            If (chunkValue Mod 10) > 0 Then
                hundredsStr = hundredsStr & "-" & units(chunkValue Mod 10)
            End If
        End Select
        
        If scaleunits(count) <> "" Then
          chunkStr = Trim(hundredsStr) & " " & scaleunits(count)
        Else
          chunkStr = Trim(hundredsStr)
        End If
        
        tempStr = Trim(chunkStr & " " & tempStr)
      End If
      
      pNumber = pNumber \ 1000
      count = count + 1
  Loop 'While pNumber > 0
    
    SpellNumber = Trim(tempStr)
    If isNegative Then
      SpellNumber = "Negative " & SpellNumber
    End If
    
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Gets the sum of cell numeric values in a range whose column is NOT
'   hidden and the cell does not contain an error.
' Parameter(s):
'   rng - The range to sum.
' Return(s): The sum of cell numeric values in a range whose column is NOT
'   hidden and the cell does not contain an error.
' Date Created: 2026-08-19
' Date Last Modified: 2026-08-31
'------------------------------------------------------------------------------'
Public Function SUM_VISIBLE_ROW(rng As Range) As Double
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "SUM_VISIBLE_ROW"
  
  Dim cell As Range
  Dim total As Double
    
  total = 0
  For Each cell In rng.Cells
    'Only sum cells whose column is NOT hidden and does NOT contain an error
    If _
      Not cell.EntireColumn.Hidden _
      And _
      Not IsError(cell.value) _
      And _
      IsNumeric(cell.value) _
      Then
        total = total + cell.value
    End If
  Next cell
    
  SUM_VISIBLE_ROW = total
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Determines if the specified range contains the specified value.
' Remarks: It is not case sensitive.
' Parameter(s):
'   rng - the Range to search.
'   val - the value to search for.
' Returns - True if the specified range contains the specified value;
'   otherwise, False.
' Date Created: 2026-08-10
' Date Last Modified: 22026-08-10
'------------------------------------------------------------------------------'
Function RngContainsVal(rng As Range, val As Variant) As Boolean
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RngContainsVal"
  
  Dim foundCell As Range
  
  Set foundCell = rng.Find( _
    What:=val, _
    LookIn:=xlValues, _
    LookAt:=xlWhole, _
    matchCase:=False)
  RngContainsVal = Not foundCell Is Nothing
    
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Creates a string representation of the specified member of the
'   ValidationSrcType enumeration.
' Parameter(s):
'   vst - A member of the ValidationSrcType enumeration.
' Returns - A string representation of the specified member of the
'   ValidationSrcType enumeration.
' Date Created: 2026-07-04
' Date Last Modified: 22026-07-12
'------------------------------------------------------------------------------'
Public Function ValSrcTypeToString(vst As ValidationSrcType) As String
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "ValSrcTypeToString"
  
  Select Case vst
    Case vstNoValidation
      ValSrcTypeToString = "No Validation"
    Case vstHardCodedValue
      ValSrcTypeToString = "Hard-Coded"
    Case vstCellRef
      ValSrcTypeToString = "Cell Ref"
    Case vstDefinedName
      ValSrcTypeToString = "Defined Name"
    Case vstFunction
      ValSrcTypeToString = "Function"
    Case Else
      ValSrcTypeToString = "Undefined"
  End Select

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Determines if the specified table the contains column (field)
'   specified by column heading.
' Parameter(s):
'   pTblLstObj - The table to search.
'   pColHeading - The column heading to search for.
'   pListCol - The ListColumn object obtained from the specified table and
'     column. If the column does not exists, it is set to Nothiong. It is an
'     optional parameter with a default value of Nothing.
' Returns - True if the specified table contains the column (field) specified
'   by column heading; Otherwise, False
' Date Created: 2026-08-15
' Date Last Modified: 22026-08-15
'------------------------------------------------------------------------------'
Public Function TblHasCol( _
  ByVal pTblLstObj As ListObject, _
  ByVal pColHeading As String, _
  Optional ByRef pListCol As ListColumn _
  ) As Boolean
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "TblHasCol"
  
  Set pListCol = Nothing
  
  If pTblLstObj Is Nothing Or Trim(pColHeading) = "" Then
    err.Raise _
      Number:=VBA_ERR_INVALID_PROCEDURE_CALL_OR_ARGUMENT, _
      Source:=MODULE_NAME & "." & METHOD_NAME, _
      Description:="Invalid table list object or empty column heading"
      GoTo Exit_Proc
  End If
  
  On Error Resume Next
  Set pListCol = pTblLstObj.ListColumns(pColHeading)
  On Error GoTo Err_Proc
    
  TblHasCol = Not pListCol Is Nothing
  
Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Creates a string representation of the specified member of the
'   VbCallType enumeration.
' Parameter(s):
'   pCallType - A member of the vbCallType enumeration.
' Return(s): A string representation of the specified member of the
'   VbCallType enumeration.
' Date Created: 2026-05-08
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Function VbCallTypeToString(pCallType As VbCallType) As String
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "VbCallTypeToString"
  
  Select Case pCallType
    Case VbGet
      VbCallTypeToString = "Get"
    Case VbLet
      VbCallTypeToString = "Let"
    Case VbMethod
      VbCallTypeToString = "Method"
    Case VbSet
      VbCallTypeToString = "Set"
    Case Else
      VbCallTypeToString = "Unknown"
  End Select

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Creates a string representation of the specified member of the
'   XlDVType enumeration (Validation.Type).
' Parameter(s):
'   dvType - A member of the XlDVType enumeration.
' Return(s): A string representation of the specified member of the
'   XlDVType enumeration (Validation.Type).
' Date Created: 2026-07-07
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Function XlDVTypeToString(dvType As XlDVType) As String
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "XlDVTypeToString"
  
  Select Case dvType
    Case xlValidateInputOnly
      XlDVTypeToString = "Input Only"
    Case xlValidateWholeNumber
      XlDVTypeToString = "Whole Number"
    Case xlValidateDecimal
      XlDVTypeToString = "Decimal"
    Case xlValidateList
      XlDVTypeToString = "List"
    Case xlValidateDate
      XlDVTypeToString = "Date"
    Case xlValidateTime
      XlDVTypeToString = "Time"
    Case xlValidateTextLength
      XlDVTypeToString = "Text Length"
    Case xlValidateCustom
      XlDVTypeToString = "Custom"
    Case Else
      XlDVTypeToString = "Undefined"
  End Select

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function

'------------------------------------------------------------------------------'
' Summary: Creates a string representation of the specified member of the
'   XlFormatConditionOperator enumeration.
' Parameter(s):
'   op - A member of the XlFormatConditionOperator enumeration.
' Return(s): A string representation of the specified member of the
'   XlFormatConditionOperator enumeration.
' Date Created: 2026-07-07
' Date Last Modified: 2026-07-12
'------------------------------------------------------------------------------'
Public Function XlFormatConditionOperatorToString( _
  op As XlFormatConditionOperator) As String
  
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "XlFormatConditionOperatorToString"
  
  Select Case op
    Case xlBetween
      XlFormatConditionOperatorToString = "Between"
    Case xlNotBetween
      XlFormatConditionOperatorToString = "Not Between"
    Case xlEqual
      XlFormatConditionOperatorToString = "Equal"
    Case xlNotEqual
      XlFormatConditionOperatorToString = "Not Equal"
    Case xlGreater
      XlFormatConditionOperatorToString = "Greater"
    Case xlLess
      XlFormatConditionOperatorToString = "Less"
    Case xlGreaterEqual
      XlFormatConditionOperatorToString = "Greater Equal"
    Case xlLessEqual
      XlFormatConditionOperatorToString = "Less Equal"
    Case Else
      XlFormatConditionOperatorToString = "Undefined"
  End Select

Exit_Proc:
  Exit Function
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Function
