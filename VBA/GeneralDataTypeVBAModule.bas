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

'   Byte Data Type
'   Byte variables are stored as single, unsigned, 8-bit (1-byte) numbers ranging in value from 0-255.
'   https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/byte-data-type

Public Const ZERO_BYTE_VAL As Byte = 0
Public Const TRUE_OFFICE_BYTE_VAL As Byte = 1
Public Const TRUE_BYTE_VAL As Byte = 1
Public Const FALSE_BYTE_VAL As Byte = ZERO_BYTE_VAL
Public Const INBETWEEN_TRUE_FALSE_BYTE_VAL As Byte = 2
Public Const NULL_BYTE_VAL As Byte = 3
'   Actual default
Public Const DEFAULT_BYTE_VAL As Byte = 0
Public Const UNASSIGNED_BYTE_VAL As Byte = 4
Public Const UNKNOWN_BYTE_VAL As Byte = 5
Public Const UNCATEGORISED_BYTE_VAL As Byte = 6
Public Const UNDEFINED_BYTE_VAL As Byte = 7
Public Const ERROR_BYTE_VAL As Byte = 8
Public Const ERROR_NAME_BYTE_VAL As Byte = 9
Public Const ERROR_NUM_BYTE_VAL As Byte = 10
Public Const ERROR_NUM_DIV_ZERO_BYTE_VAL As Byte = 11
Public Const ERROR_NUM_UNDEFINED_BYTE_VAL As Byte = 12
Public Const ERROR_REF_BYTE_VAL As Byte = 13
Public Const NOT_AVAIL_BYTE_VAL As Byte = 14
Public Const NOT_APPLICABLE_BYTE_VAL As Byte = 15
Public Const MISC_BYTE_VAL As Byte = 16
Public Const OTHER_BYTE_VAL As Byte = 17
Public Const TEST_BYTE_VAL As Byte = 18
Public Const MULTI_VAL_BYTE_VAL As Byte = 19
Public Const ALL_VAL_BYTE_VAL As Byte = 20
Public Const VALID_BYTE_VAL As Byte = 21
Public Const NEGATIVE_MAX_BYTE_VAL As Byte = 253
Public Const NEGATIVE_INFINITY_BYTE_VAL As Byte = 252
Public Const NEGATIVE_EPSILON_BYTE_VAL As Byte = 22
Public Const EPSILON_BYTE_VAL As Byte = 23
Public Const INFINITY_BYTE_VAL As Byte = 254
Public Const MAX_BYTE_VAL As Byte = 255

'   Integer Data Type
'   https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/integer-data-type
'   Signed 16-bit values that range between -32768 to 32767

Public Const ZERO_INTEGER_VAL As Integer = 0
Public Const TRUE_OFFICE_INTEGER_VAL As Integer = -1
Public Const TRUE_INTEGER_VAL As Integer = 1
Public Const FALSE_INTEGER_VAL As Integer = ZERO_INTEGER_VAL
Public Const INBETWEEN_TRUE_FALSE_INTEGER_VAL As Integer = -32745
Public Const NULL_INTEGER_VAL As Integer = -32766
'   Actual default
Public Const DEFAULT_INTEGER_VAL As Integer = 0
Public Const UNASSIGNED_INTEGER_VAL As Integer = -32765
Public Const UNKNOWN_INTEGER_VAL As Integer = -32764
Public Const UNCATEGORISED_INTEGER_VAL As Integer = -32763
Public Const UNDEFINED_INTEGER_VAL As Integer = -32762
Public Const ERROR_INTEGER_VAL As Integer = -32761
Public Const ERROR_NAME_INTEGER_VAL As Integer = -32760
Public Const ERROR_NUM_INTEGER_VAL As Integer = -32759
Public Const ERROR_NUM_DIV_ZERO_INTEGER_VAL As Integer = -32758
Public Const ERROR_NUM_UNDEFINED_INTEGER_VAL As Integer = -32757
Public Const ERROR_REF_INTEGER_VAL As Integer = -32756
Public Const NOT_AVAIL_INTEGER_VAL As Integer = -32755
Public Const NOT_APPLICABLE_INTEGER_VAL As Integer = -32754
Public Const MISC_INTEGER_VAL As Integer = -32753
Public Const OTHER_INTEGER_VAL As Integer = -32752
Public Const TEST_INTEGER_VAL As Integer = -32751
Public Const MULTI_VAL_INTEGER_VAL As Integer = -32750
Public Const ALL_VAL_INTEGER_VAL As Integer = -32749
Public Const VALID_INTEGER_VAL As Integer = -32748
Public Const NEGATIVE_MAX_INTEGER_VAL As Integer = -32768
Public Const NEGATIVE_INFINITY_INTEGER_VAL As Integer = -32767
Public Const NEGATIVE_EPSILON_INTEGER_VAL As Integer = -32747
Public Const EPSILON_INTEGER_VAL As Integer = -32746
Public Const INFINITY_INTEGER_VAL As Integer = 32766
Public Const MAX_INTEGER_VAL As Integer = 32767

'   Now that error / exception handling "codes" have taken up specific integer values
'   , those values shouldn't be usable for other means such as ordinary calculations or variable assignment
Public Const USABLE_UPPER_BOUNDARY_INTEGER_VAL As Integer = 32765
Public Const USABLE_LOWER_BOUNDARY_INTEGER_VAL As Integer = -32744

'   Long Data Type
'   Holds signed 32-bit (8-byte) integers ranging in value from -2147483648 to 2147483647
'   https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/long-data-type

Public Const ZERO_LONG_VAL As Long = 0
Public Const TRUE_OFFICE_LONG_VAL As Long = -1
Public Const TRUE_LONG_VAL As Long = 1
Public Const FALSE_LONG_VAL As Long = ZERO_LONG_VAL
Public Const INBETWEEN_TRUE_FALSE_LONG_VAL As Long = -2147483479
Public Const NULL_LONG_VAL As Long = -2147483500
'   Actual default
Public Const DEFAULT_LONG_VAL As Long = 0
Public Const UNASSIGNED_LONG_VAL As Long = -2147483499
Public Const UNKNOWN_LONG_VAL As Long = -2147483498
Public Const UNCATEGORISED_LONG_VAL As Long = -2147483497
Public Const UNDEFINED_LONG_VAL As Long = -2147483496
Public Const ERROR_LONG_VAL As Long = -2147483495
Public Const ERROR_NAME_LONG_VAL As Long = -2147483494
Public Const ERROR_NUM_LONG_VAL As Long = -2147483493
Public Const ERROR_NUM_DIV_ZERO_LONG_VAL As Long = -2147483492
Public Const ERROR_NUM_UNDEFINED_LONG_VAL As Long = -2147483491
Public Const ERROR_REF_LONG_VAL As Long = -2147483490
Public Const NOT_AVAIL_LONG_VAL As Long = -2147483489
Public Const NOT_APPLICABLE_LONG_VAL As Long = -2147483488
Public Const MISC_LONG_VAL As Long = -2147483487
Public Const OTHER_LONG_VAL As Long = -2147483486
Public Const TEST_LONG_VAL As Long = -2147483485
Public Const MULTI_VAL_LONG_VAL As Long = -2147483484
Public Const ALL_VAL_LONG_VAL As Long = -2147483483
Public Const VALID_LONG_VAL As Long = -2147483482
Public Const NEGATIVE_MAX_LONG_VAL As Long = -2147483648
Public Const NEGATIVE_INFINITY_LONG_VAL As Long = -2147483647
Public Const NEGATIVE_EPSILON_LONG_VAL As Long = -2147483481
Public Const EPSILON_LONG_VAL As Long = -2147483480
Public Const INFINITY_LONG_VAL As Long = 2147483646
Public Const MAX_LONG_VAL As Long = 2147483647

'   Now that error / exception handling "codes" have taken up specific integer values
'   , those values shouldn't be usable for other means such as ordinary calculations or variable assignment
Public Const USABLE_UPPER_BOUNDARY_LONG_VAL As Long = 2147483645
Public Const USABLE_LOWER_BOUNDARY_LONG_VAL As Long = -2147483478

'   Double Data Type
'   Stores signed IEEE 64-bit (8-byte) floating-point numbers ranging in value from:
'       -1.79769313486231E308 to -4.94065645841247E-324 for negative values
'       4.94065645841247E-324 to 1.79769313486232E308 for positive values
'   https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/double-data-type

Public Const ZERO_DOUBLE_VAL As Double = 0#
Public Const TRUE_OFFICE_DOUBLE_VAL As Double = -1#
Public Const TRUE_DOUBLE_VAL As Double = 1#
Public Const FALSE_DOUBLE_VAL As Double = ZERO_DOUBLE_VAL
Public Const INBETWEEN_TRUE_FALSE_DOUBLE_VAL As Double = VALID_DOUBLE_VAL + 1#
Public Const NULL_DOUBLE_VAL As Double = NEGATIVE_INFINITY_DOUBLE_VAL + 1#
'   Actual default
Public Const DEFAULT_DOUBLE_VAL As Double = 0
Public Const UNASSIGNED_DOUBLE_VAL As Double = NULL_DOUBLE_VAL + 1#
Public Const UNKNOWN_DOUBLE_VAL As Double = UNASSIGNED_DOUBLE_VAL + 1#
Public Const UNCATEGORISED_DOUBLE_VAL As Double = UNKNOWN_DOUBLE_VAL + 1#
Public Const UNDEFINED_DOUBLE_VAL As Double = UNCATEGORISED_DOUBLE_VAL + 1#
Public Const ERROR_DOUBLE_VAL As Double = UNDEFINED_DOUBLE_VAL + 1#
Public Const ERROR_NAME_DOUBLE_VAL As Double = ERROR_DOUBLE_VAL + 1#
Public Const ERROR_NUM_DOUBLE_VAL As Double = ERROR_NAME_DOUBLE_VAL + 1#
Public Const ERROR_NUM_DIV_ZERO_DOUBLE_VAL As Double = ERROR_NUM_DOUBLE_VAL + 1#
Public Const ERROR_NUM_UNDEFINED_DOUBLE_VAL As Double = ERROR_NUM_DIV_ZERO_DOUBLE_VAL + 1#
Public Const ERROR_REF_DOUBLE_VAL As Double = ERROR_NUM_UNDEFINED_DOUBLE_VAL + 1#
Public Const NOT_AVAIL_DOUBLE_VAL As Double = ERROR_REF_DOUBLE_VAL + 1#
Public Const NOT_APPLICABLE_DOUBLE_VAL As Double = NOT_AVAIL_DOUBLE_VAL + 1#
Public Const MISC_DOUBLE_VAL As Double = NOT_APPLICABLE_DOUBLE_VAL + 1#
Public Const OTHER_DOUBLE_VAL As Double = MISC_DOUBLE_VAL + 1#
Public Const TEST_DOUBLE_VAL As Double = OTHER_DOUBLE_VAL + 1#
Public Const MULTI_VAL_DOUBLE_VAL As Double = TEST_DOUBLE_VAL + 1#
Public Const ALL_VAL_DOUBLE_VAL As Double = MULTI_VAL_DOUBLE_VAL + 1#
Public Const VALID_DOUBLE_VAL As Double = ALL_VAL_DOUBLE_VAL + 1#
Public Const NEGATIVE_MAX_DOUBLE_VAL As Double = -1.79769313486231E308#
Public Const NEGATIVE_INFINITY_DOUBLE_VAL As Double = NEGATIVE_MAX_DOUBLE_VAL + 1#
Public Const NEGATIVE_EPSILON_DOUBLE_VAL As Double = -4.94065645841247E-324#
Public Const EPSILON_DOUBLE_VAL As Double = 4.94065645841247E-324#
Public Const INFINITY_DOUBLE_VAL As Double = MAX_DOUBLE_VAL - 1#
Public Const MAX_DOUBLE_VAL As Double = 1.79769313486232E308#

'   Now that error / exception handling "codes" have taken up specific integer values
'   , those values shouldn't be usable for other means such as ordinary calculations or variable assignment
Public Const USABLE_UPPER_BOUNDARY_DOUBLE_VAL As Double = INFINITY_DOUBLE_VAL - 1#
Public Const USABLE_LOWER_BOUNDARY_DOUBLE_VAL As Double = INBETWEEN_TRUE_FALSE_DOUBLE_VAL + 1#

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