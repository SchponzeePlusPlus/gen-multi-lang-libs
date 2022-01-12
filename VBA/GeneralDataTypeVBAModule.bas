Attribute VB_Name = "GeneralDataTypeVBAModule"
'   Eriez Magnetics Australia MS VBA
'   General Use Module (for all MS applications)
'   GeneralDataTypeVBAModule
'   Leonard Sponza
'   Last Modified 23/07/2021 15:30
'   Date Time Version 00

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

Public Const ERROR_INTEGER_VAL As Integer = -32700
Public Const NULL_INTEGER_VAL As Integer = -32699

Public Const UNASSIGNED_LONG_VAL As Long = -1999999998
Public Const NULL_LONG_VAL As Long = -1999999999
Public Const VALID_LONG_VAL As Long = -1999999995
Public Const MAX_LONG_VAL As Long = 2147483646
Public Const INFINITY_LONG_VAL As Long = MAX_LONG_VAL
Public Const NEGATIVE_MAX_LONG_VAL As Long = -2147483646
Public Const NEGAIVE_INFINITY_LONG_VAL As Long = NEGATIVE_MAX_LONG_VAL
Public Const EPSILON_LONG_VAL As Long = -1999999994
Public Const NEGATIVE_EPSILON_LONG_VAL As Long = -1999999993
Public Const MULTI_VAL_LONG_VAL As Long = -1999999992
Public Const UNKNOWN_LONG_VAL As Long = -1999999991
Public Const NOT_APPLICABLE_LONG_VAL As Long = -1999999997
Public Const NOT_AVAILABLE_LONG_VAL As Long = -1999999990
Public Const ERROR_LONG_VAL As Long = -2000000000
Public Const TEST_LONG_VAL As Long = -1999999996

Public Const UNASSIGNED_DOUBLE_VAL As Double = -9999999997#
Public Const NULL_DOUBLE_VAL As Double = -9999999998#
Public Const VALID_DOUBLE_VAL As Double = -9999999996#
Public Const MAX_DOUBLE_VAL As Double = 1.79769313486231E308#
Public Const INFINITY_DOUBLE_VAL As Double = MAX_DOUBLE_VAL - 1
Public Const NEGATIVE_MAX_DOUBLE_VAL As Double = -1.79769313486230E308#
Public Const NEGAIVE_INFINITY_DOUBLE_VAL As Double = NEGATIVE_MAX_DOUBLE_VAL + 1
Public Const EPSILON_DOUBLE_VAL As Double = 4.94065645841246E-324#
Public Const NEGATIVE_EPSILON_DOUBLE_VAL As Double = -4.94065645841246E-324#
Public Const MULTI_VAL_DOUBLE_VAL As Double = -9999999995#
Public Const UNKNOWN_DOUBLE_VAL As Double = -9999999994#
Public Const NOT_APPLICABLE_DOUBLE_VAL As Double = -9999999993#
Public Const NOT_AVAILABLE_DOUBLE_VAL As Double = -9999999992#
Public Const ERROR_DOUBLE_VAL As Double = -9999999999#
Public Const TEST_DOUBLE_VAL As Double = -9999999991#

Public Const UNASSIGNED_STRING_VAL As String = "(Unassigned)"
Public Const NULL_STRING_VAL As String = "(Null)"
Public Const ERROR_STRING_VAL As String = "(Error)"

'   #VALUE!; CVErr(xlErrValue)
Public Const ERROR_VALUE_VARIANT_VAL As Variant = "#VALUE!"

Public Enum CustomBoolean
	TRUE_CUST_BOOL = 1
	FALSE_CUST_BOOL = 2
	UNASSIGNED_CUST_BOOL = UNASSIGNED_LONG_VAL
	NULL_CUST_BOOL = NULL_LONG_VAL
	NOT_APPLICABLE_CUST_BOOL = NOT_APPLICABLE_LONG_VAL
	ERROR_CUST_BOOL = ERROR_LONG_VAL
	TEST_CUST_BOOL = TEST_LONG_VAL
End Enum

Public Enum GeneralNumVarState
	UNASSIGNED_GNVS = UNASSIGNED_LONG_VAL
	NULL_GNVS = NULL_LONG_VAL
	VALID_GNVS = VALID_LONG_VAL
	MAX_GNVS = MAX_LONG_VAL
	INFINITY_GNVS = INFINITY_LONG_VAL
	NEGATIVE_MAX_GNVS = NEGATIVE_MAX_LONG_VAL
	NEGATIVE_INFINITY_GNVS = NEGAIVE_INFINITY_LONG_VAL
	EPSILON_GNVS = EPSILON_LONG_VAL
	NEGATIVE_EPSILON_GNVS = NEGATIVE_EPSILON_LONG_VAL
	MULTI_VAL_GNVS = MULTI_VAL_LONG_VAL
	UNKNOWN_GNVS = UNKNOWN_LONG_VAL
	NOT_APPLICABLE_GNVS = NOT_APPLICABLE_LONG_VAL
	NOT_AVAILABLE_GNVS = NOT_AVAILABLE_LONG_VAL
	ERROR_GNVS = ERROR_LONG_VAL
	TEST_GNVS = TEST_LONG_VAL
End Enum

Public Type ArrayDimensionsType
	l_bnd As Long
	u_bnd As Long
	length As Long
End Type

Public Function RETURN_ERROR_INT_VAL() As Integer
	RETURN_ERROR_INT_VAL = ERROR_INTEGER_VAL
End Function

Public Function RETURN_NULL_INT_VAL() As Integer
	RETURN_NULL_INT_VAL = NULL_INTEGER_VAL
End Function

Public Function RETURN_ERROR_LONG_VAL() As Long
	RETURN_ERROR_LONG_VAL = ERROR_LONG_VAL
End Function

Public Function RETURN_NULL_LONG_VAL() As Long
	RETURN_NULL_LONG_VAL = NULL_LONG_VAL
End Function

Public Function RETURN_UNASSIGNED_DOUBLE_VAL() As Double
	RETURN_UNASSIGNED_DOUBLE_VAL = UNASSIGNED_DOUBLE_VAL
End Function

Public Function RETURN_ERROR_DOUBLE_VAL() As Double
	RETURN_ERROR_DOUBLE_VAL = ERROR_DOUBLE_VAL
End Function

Public Function RETURN_NULL_DOUBLE_VAL() As Double
	RETURN_NULL_DOUBLE_VAL = NULL_DOUBLE_VAL
End Function

Public Function RETURN_MAX_DOUBLE_VAL() As Double
	RETURN_MAX_DOUBLE_VAL = MAX_DOUBLE_VAL
End Function

Public Function RETURN_NEGATIVE_MAX_DOUBLE_VAL() As Double
	RETURN_NEGATIVE_MAX_DOUBLE_VAL = NEGATIVE_MAX_DOUBLE_VAL
End Function

Public Function RETURN_UNASSIGNED_STRING_VAL() As String
	RETURN_UNASSIGNED_STRING_VAL = UNASSIGNED_STRING_VAL
End Function

Public Function RETURN_ERROR_STRING_VAL() As String
	RETURN_ERROR_STRING_VAL = ERROR_STRING_VAL
End Function

Public Function RETURN_NULL_STRING_VAL() As String
	RETURN_NULL_STRING_VAL = NULL_STRING_VAL
End Function

Public Function ASSIGN_VARIANT_GNVS(var As Variant) As GeneralNumVarState
	Dim result As GeneralNumVarState

	Dim var_data_type As String

	var_data_type = PRINT_VARIANT_TYPENAME(var)

	If ((var_data_type = "Long") And (var = UNASSIGNED_LONG_VAL)) Then
		result = UNASSIGNED_GNVS
	ElseIf ((var_data_type = "Long") And (var = NULL_LONG_VAL)) Then
		result = NULL_GNVS
	ElseIf ((var_data_type = "Long") And (var = VALID_LONG_VAL)) Then
		result = VALID_GNVS
	ElseIf ((var_data_type = "Long") And (var = MAX_LONG_VAL)) Then
		result = MAX_GNVS
	ElseIf ((var_data_type = "Long") And (var = INFINITY_LONG_VAL)) Then
		result = INFINITY_GNVS
	ElseIf ((var_data_type = "Long") And (var = NEGATIVE_MAX_LONG_VAL)) Then
		result = NEGATIVE_MAX_GNVS
	ElseIf ((var_data_type = "Long") And (var = NEGAIVE_INFINITY_LONG_VAL)) Then
		result = NEGATIVE_INFINITY_GNVS
	ElseIf ((var_data_type = "Long") And (var = EPSILON_LONG_VAL)) Then
		result = EPSILON_GNVS
	ElseIf ((var_data_type = "Long") And (var = NEGATIVE_EPSILON_LONG_VAL)) Then
		result = NEGATIVE_EPSILON_GNVS
	ElseIf ((var_data_type = "Long") And (var = MULTI_VAL_LONG_VAL)) Then
		result = MULTI_VAL_GNVS
	ElseIf ((var_data_type = "Long") And (var = UNKNOWN_LONG_VAL)) Then
		result = UNKNOWN_GNVS
	ElseIf ((var_data_type = "Long") And (var = NOT_APPLICABLE_LONG_VAL)) Then
		result = NOT_APPLICABLE_GNVS
	ElseIf ((var_data_type = "Long") And (var = NOT_AVAILABLE_LONG_VAL)) Then
		result = NOT_AVAILABLE_GNVS
	ElseIf ((var_data_type = "Long") And (var = ERROR_LONG_VAL)) Then
		result = ERROR_GNVS
	ElseIf ((var_data_type = "Long") And (var = TEST_LONG_VAL)) Then
		result = TEST_GNVS
	ElseIf ((var_data_type = "Long") And ((var <> UNASSIGNED_LONG_VAL) And (var <> NULL_LONG_VAL) And (var <> VALID_LONG_VAL) And (var <> MAX_LONG_VAL) And (var <> INFINITY_LONG_VAL) And (var <> NEGATIVE_MAX_LONG_VAL) And (var <> NEGAIVE_INFINITY_LONG_VAL) And (var <> EPSILON_LONG_VAL) And (var <> NEGATIVE_EPSILON_LONG_VAL) And (var <> MULTI_VAL_LONG_VAL) And (var <> UNKNOWN_LONG_VAL) And (var <> NOT_APPLICABLE_LONG_VAL) And (var <> NOT_AVAILABLE_LONG_VAL) And (var <> ERROR_LONG_VAL) And (var <> TEST_LONG_VAL))) Then
		result = VALID_GNVS
	ElseIf ((var_data_type = "Double") And (var = UNASSIGNED_DOUBLE_VAL)) Then
		result = UNASSIGNED_GNVS
	ElseIf ((var_data_type = "Double") And (var = NULL_DOUBLE_VAL)) Then
		result = NULL_GNVS
	ElseIf ((var_data_type = "Double") And (var = VALID_DOUBLE_VAL)) Then
		result = VALID_GNVS
	ElseIf ((var_data_type = "Double") And (var = MAX_DOUBLE_VAL)) Then
		result = MAX_GNVS
	ElseIf ((var_data_type = "Double") And (var = INFINITY_DOUBLE_VAL)) Then
		result = INFINITY_GNVS
	ElseIf ((var_data_type = "Double") And (var = NEGATIVE_MAX_DOUBLE_VAL)) Then
		result = NEGATIVE_MAX_GNVS
	ElseIf ((var_data_type = "Double") And (var = NEGAIVE_INFINITY_DOUBLE_VAL)) Then
		result = NEGATIVE_INFINITY_GNVS
	ElseIf ((var_data_type = "Double") And (var = EPSILON_DOUBLE_VAL)) Then
		result = EPSILON_GNVS
	ElseIf ((var_data_type = "Double") And (var = NEGATIVE_EPSILON_DOUBLE_VAL)) Then
		result = NEGATIVE_EPSILON_GNVS
	ElseIf ((var_data_type = "Double") And (var = MULTI_VAL_DOUBLE_VAL)) Then
		result = MULTI_VAL_GNVS
	ElseIf ((var_data_type = "Double") And (var = UNKNOWN_DOUBLE_VAL)) Then
		result = UNKNOWN_GNVS
	ElseIf ((var_data_type = "Double") And (var = NOT_APPLICABLE_DOUBLE_VAL)) Then
		result = NOT_APPLICABLE_GNVS
	ElseIf ((var_data_type = "Double") And (var = NOT_AVAILABLE_DOUBLE_VAL)) Then
		result = NOT_AVAILABLE_GNVS
	ElseIf ((var_data_type = "Double") And (var = ERROR_DOUBLE_VAL)) Then
		result = ERROR_GNVS
	ElseIf ((var_data_type = "Double") And (var = TEST_DOUBLE_VAL)) Then
		result = TEST_GNVS
	ElseIf ((var_data_type = "Double") And ((var <> UNASSIGNED_DOUBLE_VAL) And (var <> NULL_DOUBLE_VAL) And (var <> VALID_DOUBLE_VAL) And (var <> MAX_DOUBLE_VAL) And (var <> INFINITY_DOUBLE_VAL) And (var <> NEGATIVE_MAX_DOUBLE_VAL) And (var <> NEGAIVE_INFINITY_DOUBLE_VAL) And (var <> EPSILON_DOUBLE_VAL) And (var <> NEGATIVE_EPSILON_DOUBLE_VAL) And (var <> MULTI_VAL_DOUBLE_VAL) And (var <> UNKNOWN_DOUBLE_VAL) And (var <> NOT_APPLICABLE_DOUBLE_VAL) And (var <> NOT_AVAILABLE_DOUBLE_VAL) And (var <> ERROR_DOUBLE_VAL) And (var <> TEST_DOUBLE_VAL))) Then
		result = VALID_GNVS
	Else
		result = UNKNOWN_GNVS
	End If
	
	ASSIGN_VARIANT_GNVS = result
End Function

Public Function ASSIGN_DOUBLE_VS_FROM_GNVS(input_gnvs As GeneralNumVarState) As Double
	Dim result As Double

	Select Case input_gnvs
		Case UNASSIGNED_GNVS
			result = UNASSIGNED_DOUBLE_VAL
		Case NULL_GNVS
			result = NULL_DOUBLE_VAL
		Case VALID_GNVS
			result = VALID_DOUBLE_VAL
		Case MAX_GNVS
			result = MAX_DOUBLE_VAL
		Case INFINITY_GNVS
			result = INFINITY_DOUBLE_VAL
		Case NEGATIVE_MAX_GNVS
			result = NEGATIVE_MAX_DOUBLE_VAL
		Case NEGATIVE_INFINITY_GNVS
			result = NEGAIVE_INFINITY_DOUBLE_VAL
		Case EPSILON_GNVS
			result = EPSILON_DOUBLE_VAL
		Case NEGATIVE_EPSILON_GNVS
			result = NEGATIVE_EPSILON_DOUBLE_VAL
		Case MULTI_VAL_GNVS
			result = MULTI_VAL_DOUBLE_VAL
		Case UNKNOWN_GNVS
			result = UNKNOWN_DOUBLE_VAL
		Case NOT_APPLICABLE_GNVS
			result = NOT_APPLICABLE_DOUBLE_VAL
		Case NOT_AVAILABLE_GNVS
			result = NOT_AVAILABLE_DOUBLE_VAL
		Case ERROR_GNVS
			result = ERROR_DOUBLE_VAL
		Case TEST_GNVS
			result = TEST_DOUBLE_VAL
		Case Else
			result = ERROR_DOUBLE_VAL
	End Select

	ASSIGN_DOUBLE_VS_FROM_GNVS = result

End Function

Public Function CREATE_LONG_ONE_DIM_ARRAYDIMSTYPE(input_long_od_arr() As Long) As ArrayDimensionsType
	Dim result As ArrayDimensionsType

	result.l_bnd = LBound(input_long_od_arr, 1)
	result.u_bnd = UBound(input_long_od_arr, 1)

	result.length = result.u_bnd + result.l_bnd + 1

	CREATE_LONG_ONE_DIM_ARRAYDIMSTYPE = result
End Function

Public Function CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE(input_double_od_arr() As Double) As ArrayDimensionsType
	Dim result As ArrayDimensionsType

	result.l_bnd = LBound(input_double_od_arr, 1)
	result.u_bnd = UBound(input_double_od_arr, 1)

	result.length = result.u_bnd + result.l_bnd + 1

	CREATE_DOUBLE_ONE_DIM_ARRAYDIMSTYPE = result
End Function

Public Function CREATE_STRING_ONE_DIM_ARRAYDIMSTYPE(input_string_od_arr() As String) As ArrayDimensionsType
	Dim result As ArrayDimensionsType

	result.l_bnd = LBound(input_string_od_arr, 1)
	result.u_bnd = UBound(input_string_od_arr, 1)

	result.length = result.u_bnd + result.l_bnd + 1

	CREATE_STRING_ONE_DIM_ARRAYDIMSTYPE = result
End Function

Public Function CREATE_VARIANT_ARR_ONE_DIM_ARRAYDIMSTYPE(input_variant_od_arr() As Variant) As ArrayDimensionsType
	Dim result As ArrayDimensionsType

	result.l_bnd = LBound(input_variant_od_arr, 1)
	result.u_bnd = UBound(input_variant_od_arr, 1)

	result.length = result.u_bnd + result.l_bnd + 1

	CREATE_VARIANT_ARR_ONE_DIM_ARRAYDIMSTYPE = result
End Function