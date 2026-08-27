Option Explicit

'------------------------------------------------------------------------------'
' Module Name: modGlobalsLibrary
' Summary: Declares a library global constants and variables for use with any
'   workbook.
' Date Created: 2026-06-11
' Date Last Modified: 2026-08-28
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

' Workbook Audit Report Worksheet CodeNames and Names.
Public Const WS_DATA_VAL_ALL_RPT_CODENAME = "SheetDvAllRpt"
Public Const WS_DATA_VAL_ALL_RPT_NAME = "DV All Rpt"
Public Const WS_DATA_VAL_BY_FORMULA_RPT_CODENAME = "SheetDvByForumlaRpt"
Public Const WS_DATA_VAL_BY_FORMULA_RPT_NAME = "DV by F Rpt"
Public Const WS_DEF_NAMES_ALL_RPT_CODENAME = "SheetDefNamesAllRefs"
Public Const WS_DEF_NAMES_ALL_RPT_NAME = "Def Names All Refs"
Public Const WS_DEF_NAMES_WS_REF_RPT_CODENAME = "SheetDefNamesWsRefsRpt"
Public Const WS_DEF_NAMES_WS_REF_RPT_NAME = "Def Names Ws Refs Rpt"
Public Const WS_TBLS_WS_REF_RPT_CODENAME = "SheetTblsWsRefsRpt"
Public Const WS_TBLS_WS_REF_RPT_NAME = "Tables Ws Refs Rpt"

'------------------------------------------------------------------------------'
' Excel Limits and Configuration
'------------------------------------------------------------------------------'

' Maximum length of a worksheet name. Defined by Excel.
Public Const WS_NAME_LEN_MAX As Long = 31

' Worksheet name forbidden characters, version for searching.
Public Const WS_NAME_FORBIDDEN_CHARS As String = "\/~*[]?:"

' Worksheet name forbidden characters, readable version for MsgBox prompts.
Public Const WS_NAME_FORBIDDEN_CHARS_READABLE_LIST As String = _
  "\ ~ / * [ ] ? :"

'------------------------------------------------------------------------------'
' Error Codes
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' VBA Defined Error Codes
'------------------------------------------------------------------------------'
' Error code for the VBA defined "Invalid procedure call or argument" runtime
' error.
Public Const VBA_ERR_INVALID_PROCEDURE_CALL_OR_ARGUMENT As Long = 5

' Error code for the VBA "Application-defined or object-defined error"
' runtime error.
Public Const VBA_ERR_DEFAULT_RUNTIME_APPLICATON_OBJECT_ERR As Long = 1004

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
