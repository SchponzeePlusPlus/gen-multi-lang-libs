Attribute VB_Name = "GeneralDataTypeVBAModule"
'
'	@file GeneralDataTypeVBAModule.bas
'
'	Write description of source file here for dOxygen.
'
'	System:         Independant Tool
'	Component Name: 
'  
'	Language: VBA (General Use for ~all MS Office Applications)
'
'	@brief Can use "brief" tag to explicitly generate comments for file documentation.
'	@author Leonard Sponza
'	@version 0.41.0

'	License: MIT? 
'	Licensed Material - N/A
'	Open-Source Code
'	Address:
'		
'	Author E-Mail: 
'
'	Description / Abstract: Module file
'		This file contains the defined types for Project X:
'		Notes:
'
'	Limitations: _
'	Function:
'		1) _
'	
'	Database tables used: VBA does not refer directly to DB
'	Thread Safe: No
'	Extendable: No
'	Platform Dependencies: Microsoft Excel 32-bit
'	Compiler Options: N/A
'	Change History / Revisions:
'
'	Date			Time		Author       		Description
'	----------------------------------------------------------------------------
'	
'
'
'
'	----------------------------------------------------------------------------
'
'	Requires:
'		- GeneralCalcVBAModule
'		- GeneralXLModule


'	Routine: NameOfRoutine()
'
'	Inputs:
'		@param
'		Externals:
'		Others:
'
'	Outputs:
'		Arguments:
'		Externals:
'		@return
'		@bug
'		Errors:
'
'	Routines Called:

Option Explicit

' Numbers and general enumerations
' note that some of these features may already exist in other program language inbuilt libraries (search epsilon in C, DBL_MAX, etc.)

' imaginary number component? (complex number record structure?)
'   https://www.johndcook.com/blog/2010/06/08/c-math-gotchas/
'   https://stackoverflow.com/questions/6418807/how-to-work-with-complex-numbers-in-c
'   rational / irrational component?

' Only use error and null constants below if you know that the field / variable should never expect these values (large negative values)

'   Unassigned means uninitialised, variable defined but not assigned a value
'   Null means no value or empty value
'   Valid means a valid value (makes more sense if an enumerator is used to check state of value)
'   Max means maximum possible value like infinity
'   Epsilon small positive infinitesimal quantity

'   Natural VBA Error Data Types
'   As stated in:
'   http://www.cpearson.com/excel/ReturningErrors.aspx
'   Defined in the XLCVError Enum definition:
'   xlErrDiv0 (= 2007) returns a #DIV/0! error.
'   xlErrNA (= 2042) returns a #N/A error.
'   xlErrName (= 2029) returns a #NAME? error.
'   xlErrNull (= 2000) returns a #NULL! error.
'   xlErrNum (= 2036) returns a #NUM! error.
'   xlErrRef (= 2023) returns a #REF! error.
'   xlErrValue (= 2015) returns a #VALUE! error.
'   Use CVErr() function with one of these values to return a specific Variant Error

'   Integer Data Type
'   https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/data-types/integer-data-type
'   Signed 32-bit values that range between -2147483648 to 2147483647

Public Const ZERO_INTEGER_VAL As Integer = 0
Public Const NULL_INTEGER_VAL As Integer = -2147483500
'   Actual default
Public Const DEFAULT_INTEGER_VAL As Integer = 0
Public Const UNASSIGNED_INTEGER_VAL As Integer = -2147483499
Public Const UNKNOWN_INTEGER_VAL As Integer = -2147483498
Public Const UNCATEGORISED_INTEGER_VAL As Integer = -2147483497
Public Const UNDEFINED_INTEGER_VAL As Integer = -2147483496
'   Public Const ERROR_INTEGER_VAL As Integer = -32700
Public Const ERROR_INTEGER_VAL As Integer = -2147483495
Public Const ERROR_NAME_INTEGER_VAL As Integer = -2147483494
Public Const ERROR_NUM_INTEGER_VAL As Integer = -2147483493
Public Const ERROR_NUM_DIV_ZERO_INTEGER_VAL As Integer = -2147483492
Public Const ERROR_NUM_UNDEFINED_INTEGER_VAL As Integer = -2147483491
Public Const ERROR_REF_INTEGER_VAL As Integer = -2147483490
Public Const NOT_AVAIL_INTEGER_VAL As Integer = -2147483489
Public Const NOT_APPLICABLE_INTEGER_VAL As Integer = -2147483488
Public Const MISC_INTEGER_VAL As Integer = -2147483487
Public Const OTHER_INTEGER_VAL As Integer = -2147483486
Public Const TEST_INTEGER_VAL As Integer = -2147483485
Public Const MULTI_VAL_INTEGER_VAL As Integer = -2147483484
Public Const ALL_VAL_INTEGER_VAL As Integer = -2147483483
Public Const VALID_INTEGER_VAL As Integer = -2147483482
Public Const NEGATIVE_MAX_INTEGER_VAL As Integer = -2147483648
Public Const NEGATIVE_INFINITY_INTEGER_VAL As Integer = -2147483647
Public Const NEGATIVE_EPSILON_INTEGER_VAL As Integer = -2147483481
Public Const EPSILON_INTEGER_VAL As Integer = -2147483480
Public Const INFINITY_INTEGER_VAL As Integer = 2147483646
Public Const MAX_INTEGER_VAL As Integer = 2147483647

'   Now that error / exception handling "codes" have taken up specific integer values
'   , those values shouldn't be usable for other means such as ordinary calculations or variable assignment
Public Const USABLE_UPPER_BOUNDARY_INTEGER_VAL As Integer = 2147483645
Public Const USABLE_LOWER_BOUNDARY_INTEGER_VAL As Integer = -2147483479

'   Long Data Type
'   Holds signed 64-bit (8-byte) integers ranging in value from -9223372036854775808 through 9223372036854775807 (9.2...E+18).
'   https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/data-types/long-data-type

Public Const ZERO_LONG_VAL As Long = 0
Public Const NULL_LONG_VAL As Long = -
'   Actual default
Public Const DEFAULT_LONG_VAL As Long = 0
Public Const UNASSIGNED_LONG_VAL As Long = -
Public Const UNKNOWN_LONG_VAL As Long = -
Public Const UNCATEGORISED_LONG_VAL As Long = -
Public Const UNDEFINED_LONG_VAL As Long = -
Public Const ERROR_LONG_VAL As Long = -
Public Const ERROR_NAME_LONG_VAL As Long = -
Public Const ERROR_NUM_LONG_VAL As Long = -
Public Const ERROR_NUM_DIV_ZERO_LONG_VAL As Long = -
Public Const ERROR_NUM_UNDEFINED_LONG_VAL As Long = -
Public Const ERROR_REF_LONG_VAL As Long = -
Public Const NOT_AVAIL_LONG_VAL As Long = -
Public Const NOT_APPLICABLE_LONG_VAL As Long = -
Public Const MISC_LONG_VAL As Long = -
Public Const OTHER_LONG_VAL As Long = -
Public Const TEST_LONG_VAL As Long = -
Public Const MULTI_VAL_LONG_VAL As Long = -
Public Const ALL_VAL_LONG_VAL As Long = -
Public Const VALID_LONG_VAL As Long = -
Public Const NEGATIVE_MAX_LONG_VAL As Long = -9223372036854775808
Public Const NEGATIVE_INFINITY_LONG_VAL As Long = -9223372036854775807
Public Const NEGATIVE_EPSILON_LONG_VAL As Long = -
Public Const EPSILON_LONG_VAL As Long = -
Public Const INFINITY_LONG_VAL As Long = 9223372036854775806
Public Const MAX_LONG_VAL As Long = 9223372036854775807

'   Now that error / exception handling "codes" have taken up specific integer values
'   , those values shouldn't be usable for other means such as ordinary calculations or variable assignment
Public Const USABLE_UPPER_BOUNDARY_INTEGER_VAL As Integer = 2147483645
Public Const USABLE_LOWER_BOUNDARY_INTEGER_VAL As Integer = -2147483479

'   ULong Data Type

'   Double Data Type

Public Enum ExtendedBoolV000
    TRUE_EXT_BOOL
    FALSE_EXT_BOOL
    INBETWEEN
    NULL
    DEFAULT
    UNASSIGNED
    UNKNOWN
End Enum

Public Function PRINT_STRING_OF_VARIANT_TYPENAME_V000(var As Variant) As String
	PRINT_STRING_OF_VARIANT_TYPENAME_V000 = TypeName(var)
End Function