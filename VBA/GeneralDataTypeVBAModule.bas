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
Public Const TRUE_MS_OFFICE_LONG_VAL As Long = -1
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

Public Const ZERO_DBL_VAL As Double = 0#
Public Const TRUE_OFFICE_DBL_VAL As Double = -1#
Public Const TRUE_DBL_VAL As Double = 1#
Public Const FALSE_DBL_VAL As Double = ZERO_DBL_VAL
Public Const INBETWEEN_TRUE_FALSE_DBL_VAL As Double = VALID_DBL_VAL + 1#
Public Const NULL_DBL_VAL As Double = NEGATIVE_INFINITY_DBL_VAL + 1#
'   Actual default
Public Const DEFAULT_DBL_VAL As Double = 0
Public Const UNASSIGNED_DBL_VAL As Double = NULL_DBL_VAL + 1#
Public Const UNKNOWN_DBL_VAL As Double = UNASSIGNED_DBL_VAL + 1#
Public Const UNCATEGORISED_DBL_VAL As Double = UNKNOWN_DBL_VAL + 1#
Public Const UNDEFINED_DBL_VAL As Double = UNCATEGORISED_DBL_VAL + 1#
Public Const ERROR_DBL_VAL As Double = UNDEFINED_DBL_VAL + 1#
Public Const ERROR_NAME_DBL_VAL As Double = ERROR_DBL_VAL + 1#
Public Const ERROR_NUM_DBL_VAL As Double = ERROR_NAME_DBL_VAL + 1#
Public Const ERROR_NUM_DIV_ZERO_DBL_VAL As Double = ERROR_NUM_DBL_VAL + 1#
Public Const ERROR_NUM_UNDEFINED_DBL_VAL As Double = ERROR_NUM_DIV_ZERO_DBL_VAL + 1#
Public Const ERROR_REF_DBL_VAL As Double = ERROR_NUM_UNDEFINED_DBL_VAL + 1#
Public Const NOT_AVAIL_DBL_VAL As Double = ERROR_REF_DBL_VAL + 1#
Public Const NOT_APPLICABLE_DBL_VAL As Double = NOT_AVAIL_DBL_VAL + 1#
Public Const MISC_DBL_VAL As Double = NOT_APPLICABLE_DBL_VAL + 1#
Public Const OTHER_DBL_VAL As Double = MISC_DBL_VAL + 1#
Public Const TEST_DBL_VAL As Double = OTHER_DBL_VAL + 1#
Public Const MULTI_VAL_DBL_VAL As Double = TEST_DBL_VAL + 1#
Public Const ALL_VAL_DBL_VAL As Double = MULTI_VAL_DBL_VAL + 1#
Public Const VALID_DBL_VAL As Double = ALL_VAL_DBL_VAL + 1#
Public Const NEGATIVE_MAX_DBL_VAL As Double = -1.79769313486231E308#
Public Const NEGATIVE_INFINITY_DBL_VAL As Double = NEGATIVE_MAX_DBL_VAL + 1#
Public Const NEGATIVE_EPSILON_DBL_VAL As Double = -4.94065645841247E-324#
Public Const EPSILON_DBL_VAL As Double = 4.94065645841247E-324#
Public Const INFINITY_DBL_VAL As Double = MAX_DBL_VAL - 1#
Public Const MAX_DBL_VAL As Double = 1.79769313486232E308#

'   Now that error / exception handling "codes" have taken up specific integer values
'   , those values shouldn't be usable for other means such as ordinary calculations or variable assignment
Public Const USABLE_UPPER_BOUNDARY_DBL_VAL As Double = INFINITY_DBL_VAL - 1#
Public Const USABLE_LOWER_BOUNDARY_DBL_VAL As Double = INBETWEEN_TRUE_FALSE_DBL_VAL + 1#

'   String Data Type

Public Const ZERO_STR_VAL As String = "0"
Public Const TRUE_OFFICE_STR_VAL As String = "TRUE"
Public Const TRUE_STR_VAL As String = TRUE_OFFICE_STR_VAL
Public Const FALSE_STR_VAL As String = "FALSE"
Public Const INBETWEEN_TRUE_FALSE_STR_VAL As String = "(INBETWEEN TRUE & FALSE)"
'   xlErrNull (= 2000) returns a __ error.
Public Const NULL_STR_VAL As String = "#NULL!"
'   Actual default
Public Const DEFAULT_STR_VAL As String = "(DEFAULT)"
Public Const UNASSIGNED_STR_VAL As String = "(UNASSIGNED)"
Public Const UNKNOWN_STR_VAL As String = "(UNKNOWN)"
Public Const UNCATEGORISED_STR_VAL As String = "(UNCATEGORISED)"
Public Const UNDEFINED_STR_VAL As String = "(UNDEFINED)"
'   xlErrValue (= 2015) returns a #VALUE! error. (NOT USED)
Public Const ERROR_STR_VAL As String = "(ERROR)"
'   xlErrName (= 2029) returns a __ error.
Public Const ERROR_NAME_STR_VAL As String = "#NAME?"
'   xlErrNum (= 2036) returns a __ error.
Public Const ERROR_NUM_STR_VAL As String = "#NUM!"
'   xlErrDiv0 (= 2007) returns a __ error.
Public Const ERROR_NUM_DIV_ZERO_STR_VAL As String = "#DIV/0!"
Public Const ERROR_NUM_UNDEFINED_STR_VAL As String = "(ERROR_NUM_UNDEF)"
'   xlErrRef (= 2023) returns a __ error.
Public Const ERROR_REF_STR_VAL As String = "#REF!"
'   xlErrNA (= 2042) returns a __ error.
Public Const NOT_AVAIL_STR_VAL As String = "#N/A"
Public Const NOT_APPLICABLE_STR_VAL As String = "(NOT_APPLICABLE)"
Public Const MISC_STR_VAL As String = "(MISCELLANEOUS)"
Public Const OTHER_STR_VAL As String = "(OTHER)"
Public Const TEST_STR_VAL As String = "(TEST)"
Public Const MULTI_VAL_STR_VAL As String = "(MULTI-VALUE)"
Public Const ALL_VAL_STR_VAL As String = "(ALL-VALUES)"
Public Const VALID_STR_VAL As String = "(VALID)"

'   Numeric constants don't really apply to String
' Public Const NEGATIVE_MAX_STR_VAL As String = -1.79769313486231E308#
' Public Const NEGATIVE_INFINITY_STR_VAL As String = NEGATIVE_MAX_STR_VAL + 1#
' Public Const NEGATIVE_EPSILON_STR_VAL As String = -4.94065645841247E-324#
' Public Const EPSILON_STR_VAL As String = 4.94065645841247E-324#
' Public Const INFINITY_STR_VAL As String = MAX_STR_VAL - 1#
' Public Const MAX_STR_VAL As String = 1.79769313486232E308#

'   CVErr(xlErrValue)
Public Const ERROR_VALUE_VARIANT_VAL As Variant = "#VALUE!"

'   Extended Boolean Enumerator V000
Public Enum ExtBoolEnumV000
    TRUE_EBE = TRUE_MS_OFFICE_LONG_VAL
    FALSE_EBE = FALSE_LONG_VAL
    INBETWEEN_T_F_EBE = INBETWEEN_TRUE_FALSE_LONG_VAL
    NULL_EBE = NULL_LONG_VAL
    '   DEFAULT_LONG_VAL = 0 so can't be used
    DEFAULT_EBE = 2
    UNASSIGNED_EBE = UNASSIGNED_LONG_VAL
    UNKNOWN_EBE = UNKNOWN_LONG_VAL
    NOT_APPLICABLE_EBE = NOT_APPLICABLE_LONG_VAL
    ERROR_EBE = ERROR_LONG_VAL
    TEST_EBE = TEST_LONG_VAL
End Enum

'   ExtendedExceptionHandlingNumericStatesEnum
'   EEHNSE
Public Enum ExtExcptnHandlngNumStatesEnum
    UNASSIGNED_EEHNSE = UNASSIGNED_LONG_VAL
    NULL_EEHNSE = NULL_LONG_VAL
    VALID_EEHNSE = VALID_LONG_VAL
    MAX_EEHNSE = MAX_LONG_VAL
    INFINITY_EEHNSE = INFINITY_LONG_VAL
    NEGATIVE_MAX_EEHNSE = NEGATIVE_MAX_LONG_VAL
    NEGATIVE_INFINITY_EEHNSE = NEGATIVE_INFINITY_LONG_VAL
    EPSILON_EEHNSE = EPSILON_LONG_VAL
    NEGATIVE_EPSILON_EEHNSE = NEGATIVE_EPSILON_LONG_VAL
    MULTI_VAL_EEHNSE = MULTI_VAL_LONG_VAL
    UNKNOWN_EEHNSE = UNKNOWN_LONG_VAL
    NOT_APPLICABLE_EEHNSE = NOT_APPLICABLE_LONG_VAL
    NOT_AVAIL_EEHNSE = NOT_AVAIL_LONG_VAL
    ERROR_EEHNSE = ERROR_LONG_VAL
    TEST_EEHNSE = TEST_LONG_VAL
End Enum

'   E(xtended ?) E(rror / xception ?) H(andling ?) Long Data Type
Public Type EehnsLongTypeV000
    byteEehns As Byte
    val As Long
End Type

Public Type EehnsVariantTypeV000
    byteEehns As Byte
    val As Variant
End Type

'   Array Dimension Proportions Type V000
Public Type ArrDimPropsTypeV000
    dim_bndry_lwr As Long
    dim_bndry_upr As Long
    dim_length As Long
End Type

Public Function RETURN_ERROR_INT_VAL_V000() As Integer
    RETURN_ERROR_INT_VAL_V000 = ERROR_INTEGER_VAL
End Function

Public Function RETURN_NULL_INT_VAL_V000() As Integer
    RETURN_NULL_INT_VAL_V000 = NULL_INTEGER_VAL
End Function

Public Function RETURN_ERROR_LONG_VAL_V000() As Long
    RETURN_ERROR_LONG_VAL_V000 = ERROR_LONG_VAL
End Function

Public Function RETURN_NULL_LONG_VAL_V000() As Long
    RETURN_NULL_LONG_VAL_V000 = NULL_LONG_VAL
End Function

Public Function RETURN_UNASSIGNED_DBL_VAL_V000() As Double
    RETURN_UNASSIGNED_DBL_VAL_V000 = UNASSIGNED_DBL_VAL
End Function

Public Function RETURN_ERROR_DBL_VAL_V000() As Double
    RETURN_ERROR_DBL_VAL_V000 = ERROR_DBL_VAL
End Function

Public Function RETURN_NULL_DBL_VAL_V000() As Double
    RETURN_NULL_DBL_VAL_V000 = NULL_DBL_VAL
End Function

Public Function RETURN_MAX_DBL_VAL_V000() As Double
    RETURN_MAX_DBL_VAL_V000 = MAX_DBL_VAL
End Function

Public Function RETURN_NEGATIVE_MAX_DBL_VAL_V000() As Double
    RETURN_NEGATIVE_MAX_DBL_VAL_V000 = NEGATIVE_MAX_DBL_VAL
End Function

Public Function RETURN_UNASSIGNED_STR_VAL_V000() As String
    RETURN_UNASSIGNED_STR_VAL_V000 = UNASSIGNED_STR_VAL
End Function

Public Function RETURN_ERROR_STR_VAL_V000() As String
    RETURN_ERROR_STR_VAL_V000 = ERROR_STR_VAL
End Function

Public Function RETURN_NULL_STR_VAL_V000() As String
    RETURN_NULL_STR_VAL_V000 = NULL_STR_VAL
End Function

Public Function PRINT_STR_OF_VARIANT_TYPENAME_V000(var As Variant) As String
	PRINT_STR_OF_VARIANT_TYPENAME_V000 = TypeName(var)
End Function

Public Function ASSIGN_LONG_VAL_BYTE_EEHNS(in_long_val As Long) As Byte
    Dim result As Byte

    result = UNASSIGNED_BYTE_VAL
    
    ' If (in_long_val = UNASSIGNED_LONG_VAL) Then
    ' Else
    ' End If

    ASSIGN_LONG_VAL_BYTE_EEHNS = result
End Function

Public Function ASSIGN_VARIANT_VAL_BYTE_EEHNS(in_var As Variant) As Byte
    Dim result As Byte
    Dim var_data_type As String

    result = UNASSIGNED_BYTE_VAL

    '   var_data_type = PRINT_STR_OF_VARIANT_TYPENAME_V000(in_var)
    var_data_type = TypeName(in_var)

    ASSIGN_VARIANT_VAL_BYTE_EEHNS = result
End Function

Public Function ASSIGN_LONG_VAL_BYTE_EEHNSE(in_long_val As Long) As ExtExcptnHandlngNumStatesEnum
    Dim result As ExtExcptnHandlngNumStatesEnum

    result = UNASSIGNED_BYTE_VAL
    
    If (in_long_val = UNASSIGNED_LONG_VAL) Then
        result = UNASSIGNED_EEHNSE
    ElseIf (in_long_val = NULL_LONG_VAL) Then
        result = NULL_EEHNSE
    ElseIf (in_long_val = VALID_LONG_VAL) Then
        result = VALID_EEHNSE
    ElseIf (in_long_val = MAX_LONG_VAL) Then
        result = MAX_EEHNSE
    ElseIf (in_long_val = INFINITY_LONG_VAL) Then
        result = INFINITY_EEHNSE
    ElseIf (in_long_val = _LONG_VAL) Then
        result = _EEHNSE
    ElseIf (in_long_val = _LONG_VAL) Then
        result = _EEHNSE
    ElseIf (in_long_val = _LONG_VAL) Then
        result = _EEHNSE
    ElseIf (in_long_val = _LONG_VAL) Then
        result = _EEHNSE
    ElseIf (in_long_val = _LONG_VAL) Then
        result = _EEHNSE
    ElseIf (in_long_val = _LONG_VAL) Then
        result = _EEHNSE
    ElseIf (in_long_val = _LONG_VAL) Then
        result = _EEHNSE
    ElseIf (in_long_val = _LONG_VAL) Then
        result = _EEHNSE
    ElseIf (in_long_val = _LONG_VAL) Then
        result = _EEHNSE
    ElseIf (in_long_val = _LONG_VAL) Then
        result = _EEHNSE
    ElseIf (in_long_val = _LONG_VAL) Then
        result = _EEHNSE
    Else
    End If

    ASSIGN_LONG_VAL_BYTE_EEHNSE = result
End Function

Public Function ASSIGN_VARIANT_VAL_BYTE_EEHNSE(in_var As Variant) As ExtExcptnHandlngNumStatesEnum
    Dim result As ExtExcptnHandlngNumStatesEnum
    Dim var_data_type As String

    result = UNASSIGNED_BYTE_VAL

    '   var_data_type = PRINT_STR_OF_VARIANT_TYPENAME_V000(in_var)
    var_data_type = TypeName(in_var)

    ASSIGN_VARIANT_VAL_BYTE_EEHNSE = result
End Function

'   Need to enter cast functions