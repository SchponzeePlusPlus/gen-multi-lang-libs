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
Public Const UNASSIGNED_INTEGER_VAL As Integer = 
Public Const UNKNOWN_INTEGER_VAL As Integer = 
Public Const UNCATEGORISED_INTEGER_VAL As Integer = 
Public Const UNDEFINED_INTEGER_VAL As Integer = 
Public Const ERROR_INTEGER_VAL As Integer = -32700
Public Const ERROR_NAME_INTEGER_VAL As Integer = 
Public Const ERROR_NUM_INTEGER_VAL As Integer = 
Public Const ERROR_NUM_DIV_ZERO_INTEGER_VAL As Integer = 
Public Const ERROR_NUM_UNDEFINED_INTEGER_VAL As Integer = 
Public Const ERROR_REF_INTEGER_VAL As Integer = 
Public Const NOT_AVAIL_INTEGER_VAL As Integer = 
Public Const NOT_APPLICABLE_INTEGER_VAL As Integer = 
Public Const MISC_INTEGER_VAL As Integer = 
Public Const OTHER_INTEGER_VAL As Integer = 
Public Const TEST_INTEGER_VAL As Integer = 
Public Const MULTI_VAL_INTEGER_VAL As Integer = 
Public Const ALL_VAL_INTEGER_VAL As Integer = 
Public Const VALID_INTEGER_VAL As Integer = 
Public Const NEGATIVE_MAX_INTEGER_VAL As Integer = -2147483648
Public Const NEGATIVE_INFINITY_INTEGER_VAL As Integer = -2147483647
Public Const NEGATIVE_EPSILON_INTEGER_VAL As Integer = 
Public Const EPSILON_INTEGER_VAL As Integer = 
Public Const INFINITY_INTEGER_VAL As Integer = 2147483646
Public Const MAX_INTEGER_VAL As Integer = 2147483647

'   Now that error / exception handling "codes" have taken up specific integer values
'   , those values shouldn't be usable for other means such as ordinary calculations or variable assignment
Public Const USABLE_UPPER_BOUNDARY_INTEGER_VAL As Integer = 
Public Const USABLE_LOWER_BOUNDARY_INTEGER_VAL As Integer = 

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