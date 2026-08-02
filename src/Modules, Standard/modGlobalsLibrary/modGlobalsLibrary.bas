Option Explicit

'------------------------------------------------------------------------------'
' Module Name: modGlobalsLibrary
' Summary: Declares a library global constants and variables for use with any
'   workbook.
' Date Created: 2026-06-11
' Date Last Modified: 2026-07-18
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Fields
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Constant Fields
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Private Constant Fields
'------------------------------------------------------------------------------'

' The name of this module.
Private Const MODULE_NAME As String = "modGlobalsLibrary"

'------------------------------------------------------------------------------'
' Public Constant Fields
'------------------------------------------------------------------------------'
' Number of days in a week.
Public Const DAYS_PER_WEEK As Long = 7

'------------------------------------------------------------------------------'
' Error Codes
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' VBA Defined Error Codes
'------------------------------------------------------------------------------'
Public Const VBA_ERR_INVALID_PROCEDURE_CALL_OR_ARGUMENT As Long = 5

'------------------------------------------------------------------------------'
' Enumerations
'------------------------------------------------------------------------------'

' Represents values for the first parameter of the Aggregate Funciton.
Public Enum XlAggregateFunction
    aggAverage = 1
    aggCount = 2
    aggCountA = 3
    aggMax = 4
    aggMin = 5
    aggProduct = 6
    aggSum = 9
    aggMedian = 12
    aggLarge = 14
    aggSmall = 15
End Enum

' Represents values for the first parameter of the Aggregate Funciton.
Public Enum XlAggregateOptions
    aggIgnoreNested = 0
    aggIgnoreHiddenAndNested = 1
    aggIgnoreErrorsAndNested = 2
    aggIgnoreHiddenErrorsAndNested = 3
    aggIgnoreNothing = 4
    aggIgnoreHiddenRows = 5
    aggIgnoreErrorValues = 6
End Enum

' Represents a data validaton source type.
Public Enum ValidationSrcType
  vstUndefined = 0
  vstNoValidation = 1
  vstHardCodedValue = 2
  vstCellRef = 3
  vstDefinedName = 4
  vstFunction = 5
End Enum
